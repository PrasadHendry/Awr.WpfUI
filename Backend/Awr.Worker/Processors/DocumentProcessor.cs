using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
            string sourceFilePath = FindTemplateFile(WorkerConstants.SourceLocation, _record.AwrNo);
            if (string.IsNullOrEmpty(sourceFilePath)) throw new FileNotFoundException($"Template not found: {_record.AwrNo}");

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

                // --- STAMPING LOGIC (UPDATED) ---
                foreach (Word.Section section in doc.Sections)
                {
                    // Minimal Margins
                    section.PageSetup.HeaderDistance = 12f; // ~0.17 inch
                    section.PageSetup.FooterDistance = 12f;

                    // Header
                    var headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Text = _record.GetHeaderText();
                    headerRange.Font.Name = "Calibri";
                    headerRange.Font.Size = 8;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; // Cleaner look

                    // Footer
                    var footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Text = _record.GetFooterText();
                    footerRange.Font.Name = "Calibri";
                    footerRange.Font.Size = 8;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                // Security & Save
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

        // --- RECEIPT GENERATION (UPDATED) ---
        private void PrintReceiptTable(Word.Application wordApp)
        {
            Console.WriteLine(" > Printing Receipt...");
            Word.Document doc = wordApp.Documents.Add();

            try
            {
                var range = doc.Range();

                // Title
                range.Text = "AWR DOCUMENT ISSUANCE RECEIPT\n";
                range.Font.Name = "Calibri"; // Changed from Arial
                range.Font.Bold = 1;
                range.Font.Size = 14;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.InsertParagraphAfter();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.Font.Name = "Calibri";
                range.Font.Size = 10;
                range.Font.Bold = 0;

                Word.Table table = doc.Tables.Add(range, 9, 2); // Increased rows to 9
                table.Borders.Enable = 1;
                table.Columns[1].Width = 150;
                table.Columns[2].Width = 300;

                void AddRow(int r, string k, string v)
                {
                    table.Cell(r, 1).Range.Text = k;
                    table.Cell(r, 1).Range.Font.Bold = 1;
                    table.Cell(r, 2).Range.Text = v;
                }

                // Full Details as Requested
                AddRow(1, "Request No:", _record.RequestNo);
                AddRow(2, "Document ID (AWR):", _record.AwrNo);
                AddRow(3, "Material / Product:", _record.MaterialProduct);
                AddRow(4, "Batch No:", _record.BatchNo);
                AddRow(5, "AR No:", _record.ArNo); // New Field
                AddRow(6, "Copies Issued:", _record.QtyIssued.ToString("0"));
                AddRow(7, "Issued By (QA):", _record.IssuedByUsername); // New Field
                AddRow(8, "Received By (User):", _record.PrintedByUsername);
                AddRow(9, "Timestamp:", _record.FinalActionDateText);

                // Footer
                range = doc.Range();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
                range.Text = "\nI acknowledge receipt of the above controlled documents.\n\n\n\n";
                range.Font.Name = "Calibri";
                range.Font.Size = 10;
                range.Font.Bold = 0;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "____________________________________\nSignature & Date";
                range.Font.Bold = 1;

                doc.PrintOut(Background: false);
            }
            finally
            {
                doc.Close(false);
                Marshal.ReleaseComObject(doc);
            }
        }

        private string FindTemplateFile(string directory, string fileNameNoExt)
        {
            string pathDocx = Path.Combine(directory, fileNameNoExt + ".docx");
            if (File.Exists(pathDocx)) return pathDocx;

            string pathDoc = Path.Combine(directory, fileNameNoExt + ".doc");
            if (File.Exists(pathDoc)) return pathDoc;

            return null;
        }
    }
}