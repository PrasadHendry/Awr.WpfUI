using System.Collections.Generic;
using System.Data;
using Awr.Core.Entities;
using Awr.Core.DTOs;

namespace Awr.Core.Interfaces
{
    public interface IAwrRequestRepository
    {
        List<string> CheckIfArNumberExists(string arNo, int? excludeRequestId = null);

        // NEW: Returns the gapless string directly. No connection/transaction params needed!
        string SubmitNewAwrRequest(AwrRequestSubmissionDto requestDto, string preparedByUsername);

        List<AwrItemQueueDto> GetAllAuditItems();
        List<AwrItemQueueDto> GetAuditItemsPaged(int pageNumber, int pageSize, out int totalRecords);

        List<AwrItemQueueDto> GetItemsForIssuanceQueue();
        List<AwrItemQueueDto> GetItemsForReceiptQueue(string requesterUsername);
        List<AwrItemQueueDto> GetItemsForReturnQueue(string requesterUsername);

        void IssueItem(int itemId, decimal qtyIssued, string qaUsername);
        void ReceiveItem(int itemId, string requesterUsername);
        void ReturnItem(int itemId, string requesterUsername, string remark);
        void RejectItem(int itemId, string qaUsername, string comment);

        AwrRequest GetFullRequestById(int requestId);
        List<AwrItemQueueDto> GetMySubmittedItems(string username);
        void UpdateRequestHeaderStatus(int requestId, string newStatus);
    }
}