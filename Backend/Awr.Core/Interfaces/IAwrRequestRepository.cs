using System.Collections.Generic;
using System.Data;
using Awr.Core.Entities;
using Awr.Core.DTOs;

namespace Awr.Core.Interfaces
{
    public interface IAwrRequestRepository
    {
        string GetNextRequestNumberSequenceValue();
        List<string> CheckIfArNumberExists(string arNo);

        int SubmitNewAwrRequest(IDbConnection connection, IDbTransaction transaction, string requestNo, AwrRequestSubmissionDto requestDto, string preparedByUsername);

        void InsertAwrRequestItems(IDbConnection connection, IDbTransaction transaction, int awrRequestId, List<AwrItemSubmissionDto> items);

        List<AwrItemQueueDto> GetAllAuditItems();

        List<AwrItemQueueDto> GetItemsForIssuanceQueue();
        List<AwrItemQueueDto> GetItemsForReceiptQueue(string requesterUsername);
        List<AwrItemQueueDto> GetItemsForReturnQueue(string requesterUsername);

        void IssueItem(int itemId, decimal qtyIssued, string qaUsername);
        void ReceiveItem(int itemId, string requesterUsername);

        // FIX: Added remark parameter for Void/Return justification
        void ReturnItem(int itemId, string requesterUsername, string remark);

        void RejectItem(int itemId, string qaUsername, string comment);

        AwrRequest GetFullRequestById(int requestId);
        List<AwrItemQueueDto> GetMySubmittedItems(string username);
        void UpdateRequestHeaderStatus(int requestId, string newStatus);
    }
}