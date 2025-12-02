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
            if (_record.Mode == WorkerConstants.ModeGenerate)
            {
                GenerateSecureDocument();
            }
            else if (_record.Mode == WorkerConstants.ModePrint)
            {
                PrintSecureDocument();
            }
            else
            {
                throw new InvalidOperationException($"Unknown Mode: {_record.Mode}");
            }
        }

        // --- QA: Generate (Decrypt -> Stamp -> Encrypt) ---
        private void GenerateSecureDocument()
        {
            string sourceFilePath = FindTemplateFile(WorkerConstants.SourceLocation, _record.AwrNo);
            if (string.IsNullOrEmpty(sourceFilePath))
                throw new FileNotFoundException($"Template not found for: {_record.AwrNo} (checked .docx and .doc)");

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

                // FIX 1: Provide Password when opening Template (in case it was already secured)
                try
                {
                    doc = wordApp.Documents.Open(tempFilePath, PasswordDocument: WorkerConstants.EncryptionPassword);
                }
                catch
                {
                    // Fallback: Try opening without password (if template is clean)
                    doc = wordApp.Documents.Open(tempFilePath);
                }

                // FIX 2: Unprotect if restricted so we can edit headers
                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    try { doc.Unprotect(WorkerConstants.RestrictEditPassword); } catch { /* Ignore if already unlocked */ }
                }

                // 3. Stamp Headers/Footers
                foreach (Word.Section section in doc.Sections)
                {
                    section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = _record.GetHeaderText();
                    section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = _record.GetFooterText();
                }

                // 4. Encrypt & Save
                doc.Password = WorkerConstants.EncryptionPassword;

                // Re-apply restriction
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

        // --- QC: Print (Main Doc + Receipt Table) ---
        private void PrintSecureDocument()
        {
            string fileNameBase = $"{_record.RequestNo}_{_record.AwrNo}";
            string filePath = FindTemplateFile(WorkerConstants.FinalLocation, fileNameBase);

            if (string.IsNullOrEmpty(filePath))
                throw new FileNotFoundException($"Processed file not found: {fileNameBase}");

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);

                // 1. Print Main Document
                Console.WriteLine($" > Printing Main Doc ({_record.QtyIssued:0} Copies)...");
                doc = wordApp.Documents.Open(filePath, PasswordDocument: WorkerConstants.EncryptionPassword, ReadOnly: true);

                int copies = (int)_record.QtyIssued;
                if (copies < 1) copies = 1;

                doc.PrintOut(Background: false, Copies: copies);

                doc.Close(false);
                Marshal.ReleaseComObject(doc); doc = null;

                // 2. Print Receipt Table
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
                var range = doc.Range();

                range.Text = "AWR DOCUMENT ISSUANCE RECEIPT\n";
                range.Font.Name = "Arial";
                range.Font.Bold = 1;
                range.Font.Size = 14;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "\n";
                range.InsertParagraphAfter();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.Font.Size = 11;
                range.Font.Bold = 0;

                Word.Table table = doc.Tables.Add(range, 7, 2);
                table.Borders.Enable = 1;
                table.Columns[1].Width = 150;
                table.Columns[2].Width = 300;

                void AddRow(int r, string k, string v)
                {
                    table.Cell(r, 1).Range.Text = k;
                    table.Cell(r, 1).Range.Font.Bold = 1;
                    table.Cell(r, 2).Range.Text = v;
                }

                AddRow(1, "Request No:", _record.RequestNo);
                AddRow(2, "Document ID:", _record.AwrNo);
                AddRow(3, "Material / Product:", _record.MaterialProduct);
                AddRow(4, "Batch No:", _record.BatchNo);
                AddRow(5, "Copies Issued:", _record.QtyIssued.ToString("0"));
                AddRow(6, "Received By (User):", _record.PrintedByUsername);
                AddRow(7, "Timestamp:", _record.FinalActionDateText);

                range = doc.Range();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
                range.Text = "\nI acknowledge receipt of the above controlled documents.\n\n\n\n";
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
}7