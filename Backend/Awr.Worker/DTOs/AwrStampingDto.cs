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
        public string RequestedByUsername { get; set; }
        public string IssuedByUsername { get; set; }
        public string PrintedByUsername { get; set; }

        // Helpers
        public string FinalActionDateText => DateTime.Now.ToString(Configuration.WorkerConstants.DateTimeFormat);

        // --- UPDATED HEADER (Request Info + Stamp) ---
        public string GetHeaderText()
        {
            return $"Request No: {RequestNo}\n" +
                   $"Issued By (QA): {IssuedByUsername} on {FinalActionDateText}\n" +
                   $"Status: Approved (Issued) | Processed On: {FinalActionDateText}\n" +
                   $"CONTROLLED DOCUMENT - ISSUED COPY (S/W)";
        }

        // --- UPDATED FOOTER (Material Info) ---
        public string GetFooterText()
        {
            return $"{MaterialProduct} / {BatchNo} / {ArNo}\n" +
                   $"Qty Issued: {QtyIssued:0}\n" +
                   $"AWR No.: {AwrNo}";
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