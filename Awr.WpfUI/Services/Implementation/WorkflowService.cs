using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization; // Requires ref: System.Web.Extensions
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

        // --- Sequence Generation ---
        public string GetNextRequestNumberSequenceValue()
        {
            // Fast synchronous call is acceptable here
            return _repository.GetNextRequestNumberSequenceValue();
        }

        // --- Submission (Transactional) ---
        public async Task<string> SubmitNewRequestAsync(AwrRequestSubmissionDto requestDto, string preparedByUsername, string requestNo)
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
                            // 1. Insert Header
                            _repository.SubmitNewAwrRequest(connection, transaction, requestNo, requestDto, preparedByUsername);

                            // 2. Get Generated ID
                            const string getIdSql = "SELECT Id FROM dbo.AwrRequest WHERE RequestNo = @RequestNo";
                            int newRequestId = connection.QuerySingle<int>(getIdSql, new { RequestNo = requestNo }, transaction: transaction);

                            // 3. Insert Items
                            _repository.InsertAwrRequestItems(connection, transaction, newRequestId, requestDto.Items);

                            transaction.Commit();
                            return requestNo;
                        }
                        catch (Exception)
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

        public async Task IssueItemAsync(int itemId, string qaUsername)
        {
            await Task.Run(() =>
            {
                // 1. Trigger Worker (Generate & Encrypt)
                ProcessWorkerAction(itemId, qaUsername, WorkerConstants.ModeGenerate);
                // 2. Update DB
                _repository.IssueItem(itemId, 1, qaUsername);
            });
        }

        public async Task PrintAndReceiveItemAsync(int itemId, string qcUsername)
        {
            await Task.Run(() =>
            {
                // 1. Trigger Worker (Decrypt & Print)
                ProcessWorkerAction(itemId, qcUsername, WorkerConstants.ModePrint);
                // 2. Update DB
                _repository.ReceiveItem(itemId, qcUsername);
            });
        }

        public async Task VoidItemAsync(int itemId, string qcUsername, string remark)
        {
            // No worker action for voiding
            await Task.Run(() => _repository.ReturnItem(itemId, qcUsername, remark));
        }

        public async Task RejectItemAsync(int itemId, string qaUsername, string comment)
        {
            await Task.Run(() => _repository.RejectItem(itemId, qaUsername, comment));
        }

        // --- Worker Bridge (Private Helper) ---
        private void ProcessWorkerAction(int itemId, string username, string mode)
        {
            // 1. Fetch Data
            var request = _repository.GetFullRequestById(GetRequestIdByItemId(itemId));
            var item = request.Items.First(i => i.Id == itemId);

            // 2. Construct Payload
            var dto = new AwrStampingDto
            {
                Mode = mode,
                RequestNo = request.RequestNo,
                AwrType = request.AwrType,
                ItemId = itemId,
                MaterialProduct = item.MaterialProduct,
                BatchNo = item.BatchNo,
                AwrNo = request.AwrNo,
                IssuedByUsername = mode == WorkerConstants.ModeGenerate ? username : item.IssuedByUsername,
                PrintedByUsername = mode == WorkerConstants.ModePrint ? username : null
            };

            // 3. Serialize & Encode
            var serializer = new JavaScriptSerializer(); // Legacy compatibility
            string json = serializer.Serialize(dto);
            string base64Payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

            // 4. Launch Process
            string workerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, WorkerExeName);
            if (!File.Exists(workerPath))
                throw new FileNotFoundException($"Worker executable not found at: {workerPath}");

            var startInfo = new ProcessStartInfo
            {
                FileName = workerPath,
                Arguments = $"\"CMD\" \"{base64Payload}\"", // Arg0=Dummy, Arg1=Payload
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            using (var process = Process.Start(startInfo))
            {
                if (process == null) throw new Exception("Failed to launch Worker.");
                process.WaitForExit();
                if (process.ExitCode != WorkerConstants.SuccessExitCode)
                    throw new Exception($"Worker failed. Exit Code: {process.ExitCode}");
            }
        }

        private int GetRequestIdByItemId(int itemId)
        {
            // Helper to find parent ID. In optimized SQL, we could fetch this cheaper.
            // Using existing repository method for safety.
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