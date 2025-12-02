using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Dapper;
using Awr.Core.Entities;
using Awr.Core.DTOs;
using Awr.Core.Interfaces;
using Awr.Core.Enums;

namespace Awr.Data.Repositories
{
    public class AwrRequestRepository : IAwrRequestRepository
    {
        private readonly string _connectionString;

        public AwrRequestRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        private IDbConnection GetConnection() => new SqlConnection(_connectionString);

        public string GetNextRequestNumberSequenceValue()
        {
            const string sql = "EXEC dbo.sp_GetNextAwrRequestNumber";
            using (var connection = GetConnection())
            {
                return connection.ExecuteScalar<int>(sql).ToString();
            }
        }

        public int SubmitNewAwrRequest(IDbConnection connection, IDbTransaction transaction, string requestNo, AwrRequestSubmissionDto requestDto, string preparedByUsername)
        {
            string awrNo = requestDto.Items.First().AwrNo;
            var parameters = new { RequestNo = requestNo, AwrNo = awrNo, AwrType = requestDto.Type.ToString(), PreparedByUsername = preparedByUsername, RequestComment = requestDto.RequestComment };
            const string sql = "INSERT INTO dbo.AwrRequest (RequestNo, AwrNo, AwrType, PreparedByUsername, RequestedAt, RequestComment, CurrentStatus) OUTPUT INSERTED.Id VALUES (@RequestNo, @AwrNo, @AwrType, @PreparedByUsername, GETDATE(), @RequestComment, 'PendingIssuance');";
            return connection.ExecuteScalar<int>(sql, parameters, transaction: transaction);
        }

        public void InsertAwrRequestItems(IDbConnection connection, IDbTransaction transaction, int awrRequestId, List<AwrItemSubmissionDto> items)
        {
            if (awrRequestId <= 0 || !items.Any()) return;

            // UPDATED: Using @QtyRequired instead of hardcoded 1
            const string itemSql = @"
                INSERT INTO dbo.AwrRequestItem 
                (AwrRequestId, MaterialProduct, BatchNo, ArNo, QtyRequired, Status) 
                VALUES 
                (@AwrRequestId, @MaterialProduct, @BatchNo, @ArNo, @QtyRequired, 'PendingIssuance');";

            var itemParameters = items.Select(item => new
            {
                AwrRequestId = awrRequestId,
                item.MaterialProduct,
                item.BatchNo,
                item.ArNo,
                item.QtyRequired // Maps to DTO
            });

            connection.Execute(itemSql, itemParameters, transaction: transaction);
        }

        // --- QUEUE RETRIEVAL & AUDIT ---
        private static readonly string ItemQueueSelectSql = @"
            SELECT 
                i.Id AS ItemId, 
                i.AwrRequestId AS RequestId, 
                i.MaterialProduct, 
                i.BatchNo,
                i.ArNo,              
                i.QtyRequired, 
                CASE 
                    WHEN i.Status = 'InUse' THEN " + (int)AwrItemStatus.Received + @"
                    WHEN i.Status = 'PendingIssuance' THEN " + (int)AwrItemStatus.PendingIssuance + @"
                    WHEN i.Status = 'Issued' THEN " + (int)AwrItemStatus.Issued + @"
                    WHEN i.Status = 'Returned' THEN " + (int)AwrItemStatus.Voided + @"
                    WHEN i.Status = 'Voided' THEN " + (int)AwrItemStatus.Voided + @"
                    WHEN i.Status = 'Complete' THEN " + (int)AwrItemStatus.Complete + @"
                    WHEN i.Status = 'RejectedByQa' THEN " + (int)AwrItemStatus.RejectedByQa + @"
                    ELSE 0 
                END AS Status,
                r.RequestNo, 
                r.AwrNo, 
                r.AwrType,
                r.PreparedByUsername AS RequestedBy,
                r.RequestedAt AS RequestedAt,
                r.RequestComment,    
                i.QtyIssued,
                i.IssuedByUsername AS IssuedBy,
                i.IssuedAt,
                i.ReceivedByUsername AS ReceivedBy,
                i.ReceivedAt,
                i.ReturnedByUsername AS ReturnedBy,
                i.ReturnedAt,
                i.Remark
            FROM dbo.AwrRequestItem i
            JOIN dbo.AwrRequest r ON i.AwrRequestId = r.Id ";

        public List<AwrItemQueueDto> GetAllAuditItems()
        {
            string sql = ItemQueueSelectSql + "ORDER BY r.RequestedAt DESC, i.Id;";
            using (var connection = GetConnection()) return connection.Query<AwrItemQueueDto>(sql).ToList();
        }

        public List<AwrItemQueueDto> GetItemsForIssuanceQueue()
        {
            string sql = ItemQueueSelectSql + "WHERE i.Status = 'PendingIssuance' ORDER BY r.RequestedAt DESC, i.Id;";
            using (var connection = GetConnection()) return connection.Query<AwrItemQueueDto>(sql).ToList();
        }

        public List<AwrItemQueueDto> GetItemsForReceiptQueue(string requesterUsername)
        {
            string sql = ItemQueueSelectSql + @"WHERE i.Status = 'Issued' AND r.PreparedByUsername = @Username ORDER BY r.RequestedAt DESC, i.Id;";
            using (var connection = GetConnection()) return connection.Query<AwrItemQueueDto>(sql, new { Username = requesterUsername }).ToList();
        }

        public List<AwrItemQueueDto> GetItemsForReturnQueue(string requesterUsername)
        {
            string sql = ItemQueueSelectSql + @"WHERE (i.Status = 'InUse' OR i.Status = 'Voided') AND r.PreparedByUsername = @Username ORDER BY r.RequestedAt, i.Id;";
            using (var connection = GetConnection()) return connection.Query<AwrItemQueueDto>(sql, new { Username = requesterUsername }).ToList();
        }

        public List<AwrItemQueueDto> GetMySubmittedItems(string username)
        {
            string sql = ItemQueueSelectSql + @"WHERE r.PreparedByUsername = @Username ORDER BY r.RequestedAt DESC, i.Id;";
            using (var connection = GetConnection()) return connection.Query<AwrItemQueueDto>(sql, new { Username = username }).ToList();
        }

        // --- WORKFLOW UPDATES ---

        public void IssueItem(int itemId, decimal qtyIssued, string qaUsername)
        {
            const string sql = @"UPDATE dbo.AwrRequestItem SET QtyIssued = @Qty, IssuedByUsername = @Username, IssuedAt = GETDATE(), Status = 'Issued' WHERE Id = @ItemId AND Status = 'PendingIssuance';";
            using (var connection = GetConnection())
            {
                // Note: QtyIssued tracks actual output, usually matches Request.
                if (connection.Execute(sql, new { ItemId = itemId, Username = qaUsername, Qty = qtyIssued }) == 0) throw new InvalidOperationException("Issuance failed.");
            }
        }

        public void ReceiveItem(int itemId, string requesterUsername)
        {
            const string sql = @"UPDATE dbo.AwrRequestItem SET ReceivedByUsername = @Username, ReceivedAt = GETDATE(), Status = 'InUse' WHERE Id = @ItemId AND Status = 'Issued';";
            using (var connection = GetConnection())
            {
                if (connection.Execute(sql, new { ItemId = itemId, Username = requesterUsername }) == 0) throw new InvalidOperationException("Receipt failed.");
            }
        }

        public void ReturnItem(int itemId, string requesterUsername, string remark)
        {
            using (var connection = GetConnection())
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        const string itemSql = @"
                            UPDATE dbo.AwrRequestItem 
                            SET ReturnedByUsername = @Username, 
                                ReturnedAt = GETDATE(), 
                                Remark = @Remark, 
                                Status = 'Voided' 
                            OUTPUT INSERTED.AwrRequestId 
                            WHERE Id = @ItemId AND Status = 'Issued';";

                        int? headerId = connection.ExecuteScalar<int?>(itemSql,
                            new { ItemId = itemId, Username = requesterUsername, Remark = remark },
                            transaction: transaction);

                        if (headerId == null)
                            throw new InvalidOperationException("Void failed. Item not found or not in 'Issued' state.");

                        const string headerSql = @"UPDATE dbo.AwrRequest SET CurrentStatus = 'Voided' WHERE Id = @HeaderId;";
                        connection.Execute(headerSql, new { HeaderId = headerId }, transaction: transaction);

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        public void RejectItem(int itemId, string qaUsername, string comment)
        {
            const string sql = @"UPDATE dbo.AwrRequestItem SET Remark = @Comment, Status = 'RejectedByQa' WHERE Id = @ItemId AND Status = 'PendingIssuance';";
            using (var connection = GetConnection())
            {
                if (connection.Execute(sql, new { ItemId = itemId, Comment = comment }) == 0) throw new InvalidOperationException("Rejection failed.");
            }
        }

        public AwrRequest GetFullRequestById(int requestId)
        {
            const string sql = @"SELECT * FROM dbo.AwrRequest WHERE Id = @RequestId; SELECT * FROM dbo.AwrRequestItem WHERE AwrRequestId = @RequestId;";
            using (var connection = GetConnection())
            using (var multi = connection.QueryMultiple(sql, new { RequestId = requestId }))
            {
                var request = multi.Read<AwrRequest>().SingleOrDefault();
                if (request != null) request.Items = multi.Read<AwrRequestItem>().ToList();
                return request;
            }
        }

        public void UpdateRequestHeaderStatus(int requestId, string newStatus)
        {
            const string sql = @"UPDATE dbo.AwrRequest SET CurrentStatus = @NewStatus WHERE Id = @RequestId;";
            using (var connection = GetConnection()) connection.Execute(sql, new { RequestId = requestId, NewStatus = newStatus });
        }
    }
}