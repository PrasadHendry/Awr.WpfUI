using System;
using Awr.Core.Enums;

namespace Awr.Core.DTOs
{
    public class AwrItemQueueDto
    {
        public int ItemId { get; set; }
        public int RequestId { get; set; }
        public string RequestNo { get; set; }
        public string AwrNo { get; set; }
        public AwrType AwrType { get; set; }
        public string MaterialProduct { get; set; }
        public string BatchNo { get; set; }
        public string ArNo { get; set; }

        public decimal QtyRequired { get; set; }

        // Workflow Info
        public AwrItemStatus Status { get; set; }
        public string RequestedBy { get; set; }
        public DateTime RequestedAt { get; set; }
        public string RequestComment { get; set; }

        // Audit Fields
        public decimal? QtyIssued { get; set; }
        public string IssuedBy { get; set; }
        public DateTime? IssuedAt { get; set; }

        public string ReceivedBy { get; set; } // Printed By
        public DateTime? ReceivedAt { get; set; }

        public string ReturnedBy { get; set; } // Voided By
        public DateTime? ReturnedAt { get; set; }

        // Removed Retrieval Columns

        public string Remark { get; set; }
    }
}