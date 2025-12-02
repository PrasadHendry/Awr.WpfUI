using System;
using Awr.Core.Enums;

namespace Awr.Worker.DTOs
{
    public class AwrStampingDto
    {
        // Operation Mode
        public string Mode { get; set; }

        // Header Fields
        public string RequestNo { get; set; }
        public AwrType AwrType { get; set; }

        // Item Fields
        public int ItemId { get; set; }
        public string MaterialProduct { get; set; }
        public string BatchNo { get; set; }
        public string ArNo { get; set; } // NEW
        public string AwrNo { get; set; }

        public decimal QtyIssued { get; set; } // NEW (Decimal to match DB, cast to int for printing)

        // Audit Fields
        public string IssuedByUsername { get; set; }
        public string PrintedByUsername { get; set; }

        // Helpers
        public string FinalActionDateText => DateTime.Now.ToString(Configuration.WorkerConstants.DateTimeFormat);

        // --- NEW HEADER FORMAT ---
        // (material/product) / batch no. / AR No.
        // Qty Issued: (Qty)
        // AWR No.: ...
        // CONTROLLED DOCUMENT...
        public string GetHeaderText()
        {
            return $"{MaterialProduct} / {BatchNo} / {ArNo}\n" +
                   $"Qty Issued: {QtyIssued:0}\n" +
                   $"AWR No.: {AwrNo}\n" +
                   $"CONTROLLED DOCUMENT - ISSUED COPY (S/W)";
        }

        // --- NEW FOOTER FORMAT ---
        // Request No: ...
        // Issued By (QA): ... on ...
        // Status: Approved (Issued) | Processed On: ...
        public string GetFooterText()
        {
            return $"Request No: {RequestNo}\n" +
                   $"Issued By (QA): {IssuedByUsername} on {FinalActionDateText}\n" +
                   $"Status: Approved (Issued) | Processed On: {FinalActionDateText}";
        }

        public string GetReceiptText() =>
            $"*** AWR DOCUMENT RECEIPT ***\n\n" +
            $"Request No: {RequestNo}\n" +
            $"Document: {AwrNo}\n" +
            $"Material: {MaterialProduct}\n" +
            $"Batch: {BatchNo}\n" +
            $"Copies Printed: {QtyIssued:0}\n\n" +
            $"Received & Printed By: {PrintedByUsername}\n" +
            $"Date: {FinalActionDateText}\n\n" +
            $"[ Digital Signature Placeholder ]";
    }
}