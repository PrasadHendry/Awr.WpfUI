using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Printing;
using System.Windows.Forms;
using System.Threading;
using Awr.Core.Enums;
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

        // ==========================================
        // 1. QA GENERATION
        // ==========================================
        private void GenerateSecureDocument()
        {
            string typeSubFolder = GetSubFolderForType(_record.AwrType);
            string searchDirectory = Path.Combine(WorkerConstants.SourceRoot, typeSubFolder);
            string sourceFilePath = FindTemplateFile(searchDirectory, _record.AwrNo);

            if (string.IsNullOrEmpty(sourceFilePath))
                throw new FileNotFoundException($"Template '{_record.AwrNo}' not found.");

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

                try { doc = wordApp.Documents.Open(tempFilePath, PasswordDocument: WorkerConstants.EncryptionPassword); }
                catch { doc = wordApp.Documents.Open(tempFilePath); }

                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                    try { doc.Unprotect(WorkerConstants.RestrictEditPassword); } catch { }

                SanitizeDocumentLayout(doc);
                ApplyDocumentStamps(doc);

                doc.Password = WorkerConstants.EncryptionPassword;
                if (doc.ProtectionType == Word.WdProtectionType.wdNoProtection)
                {
                    doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, NoReset: false, Password: WorkerConstants.RestrictEditPassword);
                }

                doc.SaveAs2(finalFilePath);
            }
            finally
            {
                if (doc != null) { try { doc.Close(false); } catch { } Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { try { wordApp.Quit(); } catch { } Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
                if (File.Exists(tempFilePath)) File.Delete(tempFilePath);
            }
        }

        // ==========================================
        // 2. QC PRINTING (3-STEP SEQUENCE)
        // ==========================================
        private void PrintSecureDocument()
        {
            string fileNameBase = $"{_record.RequestNo}_{_record.AwrNo}";
            string filePath = FindTemplateFile(WorkerConstants.FinalLocation, fileNameBase);

            if (string.IsNullOrEmpty(filePath)) throw new FileNotFoundException($"File not found: {fileNameBase}");

            string selectedPrinterName = null;
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                selectedPrinterName = printDialog.PrinterSettings.PrinterName;
            }
            else
            {
                throw new Exception("Printing Cancelled by User.");
            }

            ForcePrinterSettings(selectedPrinterName);

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);
                wordApp.Options.AllowReadingMode = false;
                wordApp.ActivePrinter = selectedPrinterName;

                int copies = (int)_record.QtyIssued;
                if (copies < 1) copies = 1;

                // --- STEP 1: Main AWR Doc ---
                doc = wordApp.Documents.Open(filePath, PasswordDocument: WorkerConstants.EncryptionPassword, ReadOnly: true);
                doc.PrintOut(Background: false, Copies: copies);
                doc.Close(false); Marshal.ReleaseComObject(doc); doc = null;

                Thread.Sleep(1500);

                // --- STEP 2: Receipt (Original Layout) ---
                PrintReceiptTable(wordApp, selectedPrinterName);

                Thread.Sleep(1500);

                // --- STEP 3: ALCOA Checklist ---
                PrintAlcoaChecklist(wordApp, selectedPrinterName, copies);
            }
            finally
            {
                if (doc != null) { try { doc.Close(false); } catch { } Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
            }
        }

        private void PrintAlcoaChecklist(Word.Application wordApp, string printerName, int copies)
        {
            string checklistFileName = WorkerConstants.AlcoaChecklistPrefix + _record.AwrType.ToString();
            string sourcePath = FindTemplateFile(WorkerConstants.AwrAttachmentsLocation, checklistFileName);
            if (string.IsNullOrEmpty(sourcePath)) return;

            string extension = Path.GetExtension(sourcePath);
            string tempFilePath = Path.Combine(WorkerConstants.TempLocation, Guid.NewGuid() + "_ALCOA" + extension);
            File.Copy(sourcePath, tempFilePath, true);

            Word.Document doc = null;
            try
            {
                wordApp.ActivePrinter = printerName;
                try { doc = wordApp.Documents.Open(tempFilePath, PasswordDocument: WorkerConstants.EncryptionPassword); }
                catch { doc = wordApp.Documents.Open(tempFilePath); }

                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                    try { doc.Unprotect(WorkerConstants.RestrictEditPassword); } catch { }

                SanitizeDocumentLayout(doc);
                ApplyDocumentStamps(doc);

                doc.PrintOut(Background: false, Copies: copies);
            }
            finally
            {
                if (doc != null) { try { doc.Close(false); } catch { } Marshal.ReleaseComObject(doc); }
                if (File.Exists(tempFilePath)) try { File.Delete(tempFilePath); } catch { }
            }
        }

        private void ApplyDocumentStamps(Word.Document doc)
        {
            foreach (Word.Section section in doc.Sections)
            {
                // Header Stamp
                var headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Delete();
                Word.Table stampTable = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Tables.Add(headerRange, 1, 1);
                stampTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowRight;
                stampTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
                stampTable.Borders.Enable = 1;
                stampTable.Borders.OutsideColor = Word.WdColor.wdColorDarkBlue;
                stampTable.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;
                stampTable.Range.Text = _record.GetHeaderText();
                stampTable.Range.Font.Name = "Arial";
                stampTable.Range.Font.Size = 7;
                stampTable.Range.Font.Bold = 1;
                stampTable.Range.Font.Color = Word.WdColor.wdColorDarkBlue;
                stampTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Footer Stamp
                var footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Text = _record.GetFooterText();
                footerRange.Font.Name = "Consolas";
                footerRange.Font.Size = 7;
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                footerRange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                footerRange.Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray50;
            }
        }

        private void SanitizeDocumentLayout(Word.Document doc)
        {
            foreach (Word.Section section in doc.Sections)
            {
                section.PageSetup.TopMargin = DocumentLayout.PageMarginPt;
                section.PageSetup.BottomMargin = DocumentLayout.PageMarginPt;
                section.PageSetup.LeftMargin = DocumentLayout.PageMarginPt;
                section.PageSetup.RightMargin = DocumentLayout.PageMarginPt;
                section.PageSetup.HeaderDistance = DocumentLayout.HeaderDistPt;
                section.PageSetup.FooterDistance = DocumentLayout.FooterDistPt;

                ResizeShapesInRange(section.Range);
                foreach (Word.HeaderFooter hf in section.Headers) ResizeShapesInRange(hf.Range);
                foreach (Word.HeaderFooter hf in section.Footers) ResizeShapesInRange(hf.Range);
            }

            // --- FIX FOR BLANK PAGE ---
            // Shrink trailing paragraph marks to 1pt so they don't flow to a new page
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                string pText = para.Range.Text;
                if (pText == "\r" || pText == "\v" || string.IsNullOrWhiteSpace(pText))
                {
                    para.Range.Font.Size = 1;
                    para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                    para.Format.LineSpacing = 1f;
                    para.Format.SpaceAfter = 0;
                    para.Format.SpaceBefore = 0;
                }
            }
        }

        private void ResizeShapesInRange(Word.Range range)
        {
            foreach (Word.InlineShape shape in range.InlineShapes)
            {
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                shape.Width = DocumentLayout.TargetWidthPt;
                if (shape.Height > DocumentLayout.TargetHeightPt) shape.Height = DocumentLayout.TargetHeightPt;
                shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            for (int i = range.ShapeRange.Count; i >= 1; i--)
            {
                var shape = range.ShapeRange[i];
                try
                {
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.Width = DocumentLayout.TargetWidthPt;
                    if (shape.Height > DocumentLayout.TargetHeightPt) shape.Height = DocumentLayout.TargetHeightPt;
                    shape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    shape.Top = (float)Word.WdShapePosition.wdShapeCenter;
                    shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    shape.Left = (float)Word.WdShapePosition.wdShapeCenter;
                }
                catch { }
            }
        }

        private void PrintReceiptTable(Word.Application wordApp, string printerName)
        {
            Console.WriteLine(" > Printing Receipt...");
            wordApp.ActivePrinter = printerName;

            Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Add();

                try
                {
                    if (wordApp.ActiveWindow != null) wordApp.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                }
                catch { }

                doc.PageSetup.TopMargin = 36;
                doc.PageSetup.BottomMargin = 36;
                doc.PageSetup.LeftMargin = 36;
                doc.PageSetup.RightMargin = 36;

                var range = doc.Range();

                range.Text = "SIGMA LABORATORIES PRIVATE LIMITED\n";
                range.Font.Name = "Calibri";
                range.Font.Size = 12;
                range.Font.Bold = 1;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                range.ParagraphFormat.SpaceAfter = 0;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "PLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA - 403526\n";
                range.Font.Size = 9;
                range.Font.Bold = 0;
                range.ParagraphFormat.SpaceAfter = 10;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Text = "AWR DOCUMENT ISSUANCE RECEIPT";
                range.Font.Size = 14;
                range.Font.Bold = 1;
                range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                range.ParagraphFormat.SpaceAfter = 12;
                range.InsertParagraphAfter();

                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                Word.Table table = doc.Tables.Add(range, 10, 2);
                table.Borders.Enable = 1;
                table.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 100;

                table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[1].PreferredWidth = 35;
                table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[2].PreferredWidth = 65;

                table.Range.Font.Name = "Calibri";
                table.Range.Font.Size = 10;
                table.Range.ParagraphFormat.SpaceAfter = 3;
                table.Rows.WrapAroundText = 0;

                void AddRow(int r, string k, string v)
                {
                    table.Cell(r, 1).Range.Text = k;
                    table.Cell(r, 1).Range.Font.Bold = 1;
                    table.Cell(r, 2).Range.Text = v ?? "N/A";
                    table.Cell(r, 2).WordWrap = true;
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

                range = doc.Range();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
                range.Text = "\nI hereby acknowledge receipt of the controlled document(s) listed above.\n\n\n";
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
                if (doc != null)
                {
                    doc.Close(false);
                    Marshal.ReleaseComObject(doc);
                }
            }
        }

        private void ForcePrinterSettings(string printerName)
        {
            try
            {
                using (LocalPrintServer printServer = new LocalPrintServer())
                {
                    using (PrintQueue queue = printServer.GetPrintQueue(printerName))
                    {
                        PrintTicket userTicket = queue.UserPrintTicket;
                        if (userTicket.Duplexing.HasValue) userTicket.Duplexing = Duplexing.OneSided;
                        if (userTicket.PageMediaSize != null) userTicket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
                        queue.UserPrintTicket = userTicket;
                        queue.Commit();
                    }
                }
            }
            catch { }
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
                default: return "";
            }
        }

        private string FindTemplateFile(string directory, string fileNameNoExt)
        {
            if (!Directory.Exists(directory)) return null;
            string pathDocx = Path.Combine(directory, fileNameNoExt + ".docx");
            if (File.Exists(pathDocx)) return pathDocx;
            string pathDoc = Path.Combine(directory, fileNameNoExt + ".doc");
            return File.Exists(pathDoc) ? pathDoc : null;
        }
    }
}