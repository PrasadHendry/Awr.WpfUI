using System.Collections.Generic;
using Awr.Core.Enums;

namespace Awr.Core.DTOs
{
    // DTO for submitting a new AWR request (Header + Items)
    public class AwrRequestSubmissionDto
    {
        // Header Fields
        public AwrType Type { get; set; }
        public string RequestComment { get; set; }

        // Line Items
        public List<AwrItemSubmissionDto> Items { get; set; } = new List<AwrItemSubmissionDto>();
    }
}