using System.Collections.Generic;
using System.Threading.Tasks;
using Awr.Core.DTOs;

namespace Awr.WpfUI.Services.Interfaces
{
    public interface IWorkflowService
    {
        // --- Sequence Generation ---
        string GetNextRequestNumberSequenceValue();

        // --- Submission ---
        Task<string> SubmitNewRequestAsync(AwrRequestSubmissionDto requestDto, string preparedByUsername, string requestNo);

        // --- Queue Retrieval ---
        Task<List<AwrItemQueueDto>> GetIssuanceQueueAsync();
        Task<List<AwrItemQueueDto>> GetReceiptQueueAsync(string username);
        Task<List<AwrItemQueueDto>> GetReturnQueueAsync(string username);
        Task<List<AwrItemQueueDto>> GetAllAuditItemsAsync();
        Task<List<AwrItemQueueDto>> GetMySubmittedItemsAsync(string username);

        // --- Workflow Actions (Worker & DB Updates) ---
        Task IssueItemAsync(int itemId, string qaUsername);
        Task PrintAndReceiveItemAsync(int itemId, string qcUsername);

        // --- DB Only Actions ---
        Task VoidItemAsync(int itemId, string qcUsername, string remark);
        Task RejectItemAsync(int itemId, string qaUsername, string comment);

        // --- Helper ---
        bool IsUserQAOrAdmin(string username);
    }
}