using System;
using Awr.Core.Enums;

namespace Awr.Worker.DTOs
{
    public class AwrStampingDto
    {
        // Operation Mode (GENERATE or PRINT)
        public string Mode { get; set; }

        // Header Fields
        public string RequestNo { get; set; }
        public AwrType AwrType { get; set; }

        // Item Fields
        public int ItemId { get; set; }
        public string MaterialProduct { get; set; }
        public string BatchNo { get; set; }
        public string AwrNo { get; set; }

        // Audit Fields
        public string IssuedByUsername { get; set; }  // QA User
        public string PrintedByUsername { get; set; } // QC User (For Receipt)

        // Helpers
        public string FinalActionDateText => DateTime.Now.ToString(Configuration.WorkerConstants.DateTimeFormat);

        public string GetHeaderText() => $"{MaterialProduct} / Batch: {BatchNo}\nAWR No: {AwrNo}\nTYPE: {AwrType}";

        public string GetFooterText() =>
            $"Request No: {RequestNo} | Item ID: {ItemId}\n" +
            $"Approved By (QA): {IssuedByUsername}\n" +
            $"Generated On: {FinalActionDateText}";

        public string GetReceiptText() =>
            $"*** AWR DOCUMENT RECEIPT ***\n\n" +
            $"Request No: {RequestNo}\n" +
            $"Document: {AwrNo}\n" +
            $"Material: {MaterialProduct}\n" +
            $"Batch: {BatchNo}\n\n" +
            $"Received & Printed By: {PrintedByUsername}\n" +
            $"Date: {FinalActionDateText}\n\n" +
            $"[ Digital Signature Placeholder ]";
    }
}