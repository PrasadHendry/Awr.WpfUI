using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Awr.Core.Enums; // Required for AwrType
using Awr.Worker.Configuration;
using Awr.Worker.DTOs;
using Word = Microsoft.Office.Interop.Word;

namespace Awr.Worker.Processors
{
    public class DocumentProcessor
    {
        private readonly AwrStampingDto _record;
        private static readonly object Missing = System.Reflection.Missing.Value;

        public DocumentProcessor(AwrStampingDto record)
        {
            _record = record;
        }

        public void ProcessRequest()
        {
            if (_record.Mode == WorkerConstants.ModeGenerate) GenerateSecureDocument();
            else if (_record.Mode == WorkerConstants.ModePrint) PrintSecureDocument();
            else throw new InvalidOperationException($"Unknown Mode: {_record.Mode}");
        }

        // --- QA: Generate ---
        private void GenerateSecureDocument()
        {
            // 1. Determine Correct Source Folder
            string typeSubFolder = GetSubFolderForType(_record.AwrType);
            string searchDirectory = Path.Combine(WorkerConstants.SourceRoot, typeSubFolder);

            // 2. Find Template
            string sourceFilePath = FindTemplateFile(searchDirectory, _record.AwrNo);

            if (string.IsNullOrEmpty(sourceFilePath))
                throw new FileNotFoundException($"Template '{_record.AwrNo}' not found in folder: {typeSubFolder}");

            string extension = Path.GetExtension(sourceFilePath);
            string finalFileName = $"{_record.RequestNo}_{_record.AwrNo}{extension}";
            string finalFilePath = Path.Combine(WorkerConstants.FinalLocation, finalFileName);
            string tempFilePath = Path.Combine(WorkerConstants.TempLocation, Guid.NewGuid() + extension);

            File.Copy(sourceFilePath, tempFilePath, true);

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);

                // Open (Try Secure, then Plain)
                try { doc = wordApp.Documents.Open(tempFilePath, PasswordDocument: WorkerConstants.EncryptionPassword); }
                catch { doc = wordApp.Documents.Open(tempFilePath); }

                // Unprotect
                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    try { doc.Unprotect(WorkerConstants.RestrictEditPassword); } catch { }
                }

                // Stamp Header/Footer (Calibri 8pt)
                foreach (Word.Section section in doc.Sections)
                {
                    section.PageSetup.HeaderDistance = 24f; // 0.33 inch (Safe)
                    section.PageSetup.FooterDistance = 24f;

                    var headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Text = _record.GetHeaderText();
                    headerRange.Font.Name = "Calibri";
                    headerRange.Font.Size = 8;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                    var footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Text = _record.GetFooterText();
                    footerRange.Font.Name = "Calibri";
                    footerRange.Font.Size = 8;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                // Encrypt & Save
                doc.Password = WorkerConstants.EncryptionPassword;
                if (doc.ProtectionType == Word.WdProtectionType.wdNoProtection)
                {
                    doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, NoReset: false, Password: WorkerConstants.RestrictEditPassword);
                }

                doc.SaveAs2(finalFilePath);
                Console.WriteLine($" > Generated: {finalFileName}");
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
                if (File.Exists(tempFilePath)) File.Delete(tempFilePath);
            }
        }

        // --- QC: Print ---
        private void PrintSecureDocument()
        {
            string fileNameBase = $"{_record.RequestNo}_{_record.AwrNo}";
            string filePath = FindTemplateFile(WorkerConstants.FinalLocation, fileNameBase);

            if (string.IsNullOrEmpty(filePath)) throw new FileNotFoundException($"File not found: {fileNameBase}");

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);

                Console.WriteLine($" > Printing Main Doc ({_record.QtyIssued:0} Copies)...");
                doc = wordApp.Documents.Open(filePath, PasswordDocument: WorkerConstants.EncryptionPassword, ReadOnly: true);

                int copies = (int)_record.QtyIssued;
                if (copies < 1) copies = 1;

                doc.PrintOut(Background: false, Copies: copies);

                doc.Close(false);
                Marshal.ReleaseComObject(doc); doc = null;

                PrintReceiptTable(wordApp);
            }
            finally
            {
                if (doc != null) { try { doc.Close(false); } catch { } Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
            }
        }

        private void PrintReceiptTable(Word.Application wordApp)
        {
            Console.WriteLine(" > Printing Receipt...");
            Word.Document doc = wordApp.Documents.Add();

            try
            {
                // FIX 1: Set Narrow Margins to maximize space
                doc.PageSetup.TopMargin = 36;    // 0.5 inch
                doc.PageSetup.BottomMargin = 36;
                doc.PageSetup.LeftMargin = 36;
                doc.PageSetup.RightMargin = 36;

                var range = doc.Range();

                // 1. HEADER (Compact)
                range.Text = "SIGMA LABORATORIES PRIVATE LIMITED\n";
                range.Font.Name = "Calibri";
                range.Font.Size = 12; // Reduced from 14
                range.Font.Bold = 1;
                range.ParagraphFormat.SpaceAfter = 0; // No gap
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "PLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA - 403526\n";
                range.Font.Size = 9;  // Reduced from 10
                range.Font.Bold = 0;
                range.ParagraphFormat.SpaceAfter = 10; // Small gap before title
                range.InsertParagraphAfter();

                // 2. TITLE
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "AWR DOCUMENT ISSUANCE RECEIPT";
                range.Font.Size = 14; // Reduced from 16
                range.Font.Bold = 1;
                range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                range.ParagraphFormat.SpaceAfter = 12; // Gap before table
                range.InsertParagraphAfter();

                // Reset formatting
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                // 3. TABLE
                Word.Table table = doc.Tables.Add(range, 10, 2);
                table.Borders.Enable = 1;
                table.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;

                // Layout
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 100;
                table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[1].PreferredWidth = 30; // 30% Label
                table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[2].PreferredWidth = 70; // 70% Value

                // Compact Rows
                table.Range.Font.Name = "Calibri";
                table.Range.Font.Size = 10; // Reduced from 11
                table.Range.ParagraphFormat.SpaceAfter = 3; // Tight spacing
                table.Rows.WrapAroundText = 0;

                void AddRow(int r, string k, string v)
                {
                    table.Cell(r, 1).Range.Text = k;
                    table.Cell(r, 1).Range.Font.Bold = 1;

                    table.Cell(r, 2).Range.Text = v;
                }

                AddRow(1, "Request No:", _record.RequestNo);
                AddRow(2, "Document ID (AWR):", _record.AwrNo);
                AddRow(3, "Material / Product:", _record.MaterialProduct);
                AddRow(4, "Batch No:", _record.BatchNo);
                AddRow(5, "AR No:", _record.ArNo);
                AddRow(6, "Copies Issued:", _record.QtyIssued.ToString("0"));
                AddRow(7, "Requested By (QC):", _record.RequestedByUsername);
                AddRow(8, "Issued By (QA):", _record.IssuedByUsername);
                AddRow(9, "Received By (QC):", _record.PrintedByUsername);
                AddRow(10, "Timestamp (Print):", _record.FinalActionDateText);

                // 4. FOOTER
                range = doc.Range();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter(); // Move past table

                range.Text = "\nI hereby acknowledge receipt of the controlled documents listed above.\n\n\n";
                range.Font.Name = "Calibri";
                range.Font.Size = 9;
                range.Font.Bold = 0;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.ParagraphFormat.SpaceAfter = 0;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "____________________________________\nSignature & Date";
                range.Font.Bold = 1;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                doc.PrintOut(Background: false);
            }
            finally
            {
                doc.Close(false);
                Marshal.ReleaseComObject(doc);
            }
        }

        private string GetSubFolderForType(AwrType type)
        {
            switch (type)
            {
                case AwrType.FPS: return "FPS-IMS-AWR ISSUANCE";
                case AwrType.IMS: return "FPS-IMS-AWR ISSUANCE";
                case AwrType.MICRO: return "Micro AWR Issuance";
                case AwrType.PM: return "PM AWR Issuance";
                case AwrType.RM: return "RM AWR Issuance";
                case AwrType.STABILITY: return "Stability AWR Issuance";
                case AwrType.WATER: return "Water AWR Issuance";
                default: return ""; // Root or Unknown
            }
        }

        private string FindTemplateFile(string directory, string fileNameNoExt)
        {
            // Note: directory passed here is now SourceRoot + SubFolder
            if (!Directory.Exists(directory)) return null;

            string pathDocx = Path.Combine(directory, fileNameNoExt + ".docx");
            if (File.Exists(pathDocx)) return pathDocx;

            string pathDoc = Path.Combine(directory, fileNameNoExt + ".doc");
            if (File.Exists(pathDoc)) return pathDoc;

            return null;
        }
    }
}