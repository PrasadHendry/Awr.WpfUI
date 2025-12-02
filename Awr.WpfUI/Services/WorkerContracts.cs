using System;
using Awr.Core.Enums;

// We define these locally in the UI to match the Worker's expectation
// without needing to reference the Console Application project directly.

namespace Awr.Worker.Configuration
{
    public static class WorkerConstants
    {
        public const string ModeGenerate = "GENERATE";
        public const string ModePrint = "PRINT";
        public const int SuccessExitCode = 0;
    }
}

namespace Awr.Worker.DTOs
{
    public class AwrStampingDto
    {
        public string Mode { get; set; }
        public string RequestNo { get; set; }
        public AwrType AwrType { get; set; }
        public int ItemId { get; set; }
        public string MaterialProduct { get; set; }
        public string BatchNo { get; set; }
        public string ArNo { get; set; } // NEW
        public string AwrNo { get; set; }
        public decimal QtyIssued { get; set; } // NEW
        public string IssuedByUsername { get; set; }
        public string PrintedByUsername { get; set; }
    }
}