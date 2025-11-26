using System;
using System.Collections.Generic;
using Awr.Core.Enums;

namespace Awr.Core.Entities
{
    // Corresponds to dbo.AwrRequest (Header)
    public class AwrRequest
    {
        public int Id { get; set; }
        public string RequestNo { get; set; }

        // FIX: Added AwrNo to match the database schema (VARCHAR(50) NOT NULL UNIQUE)
        public string AwrNo { get; set; }

        public AwrType AwrType { get; set; }
        public string PreparedByUsername { get; set; }
        public DateTime RequestedAt { get; set; }
        public string RequestComment { get; set; }
        public string CurrentStatus { get; set; }

        // QA Issuance fields
        public string IssuedByUsername { get; set; }
        public DateTime? IssuedAt { get; set; }

        public List<AwrRequestItem> Items { get; set; } = new List<AwrRequestItem>();
    }
}