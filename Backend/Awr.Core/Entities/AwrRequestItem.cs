using System;
using Awr.Core.Enums;

namespace Awr.Core.Entities
{
    // Corresponds to dbo.AwrRequestItem (Line Item + Audit)
    public class AwrRequestItem
    {
        public int Id { get; set; }
        public int AwrRequestId { get; set; }

        // --- REQUEST DATA ---
        public string MaterialProduct { get; set; }
        public string BatchNo { get; set; }
        public string ArNo { get; set; }
        // FIX: Removed 'AwrNo' from here. It belongs to the AwrRequest (Header).
        public decimal QtyRequired { get; set; }

        // --- WORKFLOW STATUS ---
        public AwrItemStatus Status { get; set; }

        // --- ISSUANCE (QA Action - Step 2) ---
        public decimal? QtyIssued { get; set; }
        public string IssuedByUsername { get; set; }
        public DateTime? IssuedAt { get; set; }

        // --- RECEIVED (Requester/QC Action - Step 3) ---
        public string ReceivedByUsername { get; set; }
        public DateTime? ReceivedAt { get; set; }

        // --- RETURNED (Requester/QC Action - Step 4) ---
        public string ReturnedByUsername { get; set; }
        public DateTime? ReturnedAt { get; set; }

        public string Remark { get; set; }
    }
}