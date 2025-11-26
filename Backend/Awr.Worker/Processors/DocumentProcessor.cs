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

        // --- QA: Generate (Stamp & Encrypt) ---
        private void GenerateSecureDocument()
        {
            // Input: Template (e.g., AWR-RM-001.docx)
            string sourceFilePath = FindFileByName(WorkerConstants.SourceLocation, _record.AwrNo + ".docx");
            if (string.IsNullOrEmpty(sourceFilePath)) throw new FileNotFoundException($"Template not found: {_record.AwrNo}");

            // Output: Request-Specific File (e.g., REQ-101_AWR-RM-001.docx)
            string finalFileName = $"{_record.RequestNo}_{_record.AwrNo}.docx";
            string finalFilePath = Path.Combine(WorkerConstants.FinalLocation, finalFileName);
            string tempFilePath = Path.Combine(WorkerConstants.TempLocation, Guid.NewGuid() + ".docx");

            File.Copy(sourceFilePath, tempFilePath, true);

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);

                doc = wordApp.Documents.Open(tempFilePath);

                // 1. Stamp
                foreach (Word.Section section in doc.Sections)
                {
                    section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = _record.GetHeaderText();
                    section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = _record.GetFooterText();
                }

                // 2. Encrypt & Save
                doc.Password = WorkerConstants.EncryptionPassword;
                doc.SaveAs2(finalFilePath);

                Console.WriteLine($" > Document Generated: {finalFileName}");
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
                if (File.Exists(tempFilePath)) File.Delete(tempFilePath);
            }
        }

        // --- QC: Print (Open Encrypted -> Print -> Print Receipt) ---
        private void PrintSecureDocument()
        {
            string fileName = $"{_record.RequestNo}_{_record.AwrNo}.docx";
            string filePath = Path.Combine(WorkerConstants.FinalLocation, fileName);

            if (!File.Exists(filePath)) throw new FileNotFoundException($"Processed document not found: {fileName}");

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);

                // 1. Print Main Document (Open with Password)
                Console.WriteLine(" > Printing Main Document...");
                doc = wordApp.Documents.Open(filePath, PasswordDocument: WorkerConstants.EncryptionPassword, ReadOnly: true);
                doc.PrintOut(Background: false); // Wait for print to spool
                doc.Close(false);
                Marshal.ReleaseComObject(doc); doc = null;

                // 2. Print Receipt Page
                Console.WriteLine(" > Printing Receipt...");
                doc = wordApp.Documents.Add();
                doc.Range().Text = _record.GetReceiptText();
                doc.PrintOut(Background: false);
                doc.Close(false);
            }
            finally
            {
                if (doc != null) { try { doc.Close(false); } catch { } Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
            }
        }

        private string FindFileByName(string directory, string fileName)
        {
            string path = Path.Combine(directory, fileName);
            return File.Exists(path) ? path : null;
        }
    }
}