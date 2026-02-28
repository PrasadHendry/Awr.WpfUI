using System.Collections.Generic;
using System.Threading.Tasks;
using Awr.Core.DTOs;

namespace Awr.WpfUI.Services.Interfaces
{
    public interface IWorkflowService
    {
        // --- Duplicate Check ---
        List<string> CheckIfArNumberExists(string arNo, int? excludeRequestId = null);

        List<AwrItemQueueDto> GetAuditItemsPaged(int page, int size, out int total);

        // --- Submission (NEW GAPLESS SIGNATURE) ---
        // We no longer pass the 'requestNo' from the UI. The Database generates it and returns it.
        Task<string> SubmitNewRequestAsync(AwrRequestSubmissionDto requestDto, string preparedByUsername);

        // --- Queue Retrieval ---
        Task<List<AwrItemQueueDto>> GetIssuanceQueueAsync();
        Task<List<AwrItemQueueDto>> GetReceiptQueueAsync(string username);
        Task<List<AwrItemQueueDto>> GetReturnQueueAsync(string username);
        Task<List<AwrItemQueueDto>> GetAllAuditItemsAsync();
        Task<List<AwrItemQueueDto>> GetMySubmittedItemsAsync(string username);

        // --- Workflow Actions (Worker & DB Updates) ---
        Task IssueItemAsync(int itemId, decimal qtyIssued, string qaUsername);
        Task PrintAndReceiveItemAsync(int itemId, string qcUsername);

        // --- DB Only Actions ---
        Task VoidItemAsync(int itemId, string qcUsername, string remark);
        Task RejectItemAsync(int itemId, string qaUsername, string comment);

        // --- Helper ---
        bool IsUserQAOrAdmin(string username);
    }
}