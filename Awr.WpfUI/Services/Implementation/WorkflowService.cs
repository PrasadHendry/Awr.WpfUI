using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Awr.Core.DTOs;
using Awr.Core.Interfaces;
using Awr.Data.Repositories;
using Awr.Worker.Configuration;
using Awr.Worker.DTOs;
using Awr.WpfUI.Services.Interfaces;
using Dapper;

namespace Awr.WpfUI.Services.Implementation
{
    public class WorkflowService : IWorkflowService
    {
        private readonly IAwrRequestRepository _repository;
        private readonly string _connectionString;
        private const string WorkerExeName = "Awr.Worker.exe";

        public WorkflowService()
        {
            _connectionString = ConfigurationManager.ConnectionStrings["AwrDbConnection"]?.ConnectionString;
            if (string.IsNullOrEmpty(_connectionString))
                throw new InvalidOperationException("Connection string 'AwrDbConnection' is missing.");

            _repository = new AwrRequestRepository(_connectionString);
        }

        public List<AwrItemQueueDto> GetAuditItemsPaged(int page, int size, out int total)
        {
            return _repository.GetAuditItemsPaged(page, size, out total);
        }

        // --- Sequence Generation ---
        public string GetNextRequestNumberSequenceValue()
        {
            return _repository.GetNextRequestNumberSequenceValue();
        }
        public List<string> CheckIfArNumberExists(string arNo, int? excludeRequestId = null)
        {
            return _repository.CheckIfArNumberExists(arNo, excludeRequestId);
        }

        // --- Submission ---
        public async Task<string> SubmitNewRequestAsync(AwrRequestSubmissionDto requestDto, string preparedByUsername, string requestNoPlaceholder)
        {
            return await Task.Run(() =>
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            // 1. GENERATE ID (Inside Transaction)
                            string seq = _repository.GetNextRequestNumberSequenceValue();
                            string finalRequestNo = $"AWR-{DateTime.Now:yyyyMMdd}-{seq}";

                            // 2. Insert Header
                            _repository.SubmitNewAwrRequest(connection, transaction, finalRequestNo, requestDto, preparedByUsername);

                            // 3. Get New ID
                            const string getIdSql = "SELECT Id FROM dbo.AwrRequest WHERE RequestNo = @RequestNo";
                            int newRequestId = connection.QuerySingle<int>(getIdSql, new { RequestNo = finalRequestNo }, transaction: transaction);

                            // 4. Insert Items
                            _repository.InsertAwrRequestItems(connection, transaction, newRequestId, requestDto.Items);

                            transaction.Commit();
                            return finalRequestNo; // Return the ACTUAL ID generated
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            });
        }

        // --- Queue Retrieval ---
        public async Task<List<AwrItemQueueDto>> GetIssuanceQueueAsync() =>
            await Task.Run(() => _repository.GetItemsForIssuanceQueue());

        public async Task<List<AwrItemQueueDto>> GetReceiptQueueAsync(string username) =>
            await Task.Run(() => _repository.GetItemsForReceiptQueue(username));

        public async Task<List<AwrItemQueueDto>> GetReturnQueueAsync(string username) =>
            await Task.Run(() => _repository.GetItemsForReturnQueue(username));

        public async Task<List<AwrItemQueueDto>> GetAllAuditItemsAsync() =>
            await Task.Run(() => _repository.GetAllAuditItems());

        public async Task<List<AwrItemQueueDto>> GetMySubmittedItemsAsync(string username) =>
            await Task.Run(() => _repository.GetMySubmittedItems(username));

        // --- Workflow Actions ---

        public async Task IssueItemAsync(int itemId, decimal qtyIssued, string qaUsername)
        {
            await Task.Run(() =>
            {
                // 1. Trigger Worker (Generate)
                ProcessWorkerAction(itemId, qaUsername, WorkerConstants.ModeGenerate);

                // 2. Update DB
                _repository.IssueItem(itemId, qtyIssued, qaUsername);
            });
        }

        public async Task PrintAndReceiveItemAsync(int itemId, string qcUsername)
        {
            await Task.Run(() =>
            {
                // 1. Trigger Worker (Print)
                // If the user Cancels the Print Dialog, this method will THROW an exception.
                // This stops the code here, so step 2 (Update DB) never happens.
                ProcessWorkerAction(itemId, qcUsername, WorkerConstants.ModePrint);

                // 2. Update DB (Only happens if Print was successful)
                _repository.ReceiveItem(itemId, qcUsername);
            });
        }

        public async Task VoidItemAsync(int itemId, string qcUsername, string remark)
        {
            await Task.Run(() => _repository.ReturnItem(itemId, qcUsername, remark));
        }

        public async Task RejectItemAsync(int itemId, string qaUsername, string comment)
        {
            await Task.Run(() => _repository.RejectItem(itemId, qaUsername, comment));
        }

        // --- Worker Bridge ---
        private void ProcessWorkerAction(int itemId, string username, string mode)
        {
            var request = _repository.GetFullRequestById(GetRequestIdByItemId(itemId));
            var item = request.Items.First(i => i.Id == itemId);

            var dto = new AwrStampingDto
            {
                Mode = mode,
                RequestNo = request.RequestNo,
                AwrType = request.AwrType,
                ItemId = itemId,
                MaterialProduct = item.MaterialProduct,
                BatchNo = item.BatchNo,
                ArNo = item.ArNo,
                AwrNo = request.AwrNo,
                RequestedByUsername = request.PreparedByUsername,

                // NEW: Use QtyRequired (which is the approved qty) for the stamp/print logic
                // If we are printing, QtyIssued in DB might be null if legacy, so fallback to QtyRequired
                QtyIssued = item.QtyIssued.HasValue && item.QtyIssued > 0 ? item.QtyIssued.Value : item.QtyRequired,

                IssuedByUsername = mode == WorkerConstants.ModeGenerate ? username : item.IssuedByUsername,
                PrintedByUsername = mode == WorkerConstants.ModePrint ? username : null
            };

            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(dto);
            string base64Payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

            string workerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, WorkerExeName);
            if (!File.Exists(workerPath))
                throw new FileNotFoundException($"Worker executable not found at: {workerPath}");

            var startInfo = new ProcessStartInfo
            {
                FileName = workerPath,
                Arguments = $"\"CMD\" \"{base64Payload}\"",
                UseShellExecute = false,

                // SHOW CONSOLE explicitly
                CreateNoWindow = false,
                WindowStyle = ProcessWindowStyle.Normal, // Ensure it's visible

                // Do NOT redirect output if you want the user to see the live console window.
                RedirectStandardOutput = false
            };

            using (var process = Process.Start(startInfo))
            {
                if (process == null) throw new Exception("Failed to launch Worker.");

                process.WaitForExit();

                // --- UPDATED ERROR HANDLING LOGIC ---
                if (process.ExitCode != WorkerConstants.SuccessExitCode)
                {
                    // If we are in Print Mode and it failed, it's 99% likely the user clicked Cancel on the Dialog.
                    // We throw a specific message so the UI can look cleaner.
                    if (mode == WorkerConstants.ModePrint)
                    {
                        throw new Exception("Printing Cancelled by User.");
                    }

                    // Otherwise, it's a real crash/error
                    throw new Exception($"Worker failed. Exit Code: {process.ExitCode}");
                }
            }
        }

        private int GetRequestIdByItemId(int itemId)
        {
            var r = _repository.GetFullRequestById(itemId);
            return r?.Id ?? 0;
        }

        public bool IsUserQAOrAdmin(string username)
        {
            return username.Equals("QA", StringComparison.OrdinalIgnoreCase) ||
                   username.Equals("Admin", StringComparison.OrdinalIgnoreCase);
        }
    }
}