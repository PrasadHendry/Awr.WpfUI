using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

        // --- SPECIFICATIONS (From Screenshot) ---
        // 1 cm = 28.3465 points
        private const float PageMarginPt = 14.45f;     // 0.51 cm
        private const float HeaderDistPt = 28.35f;     // 1 cm
        private const float FooterDistPt = 28.35f;     // 1 cm

        private const float TargetHeightPt = 708.5f;   // 24.99 cm
        private const float TargetWidthPt = 501f;      // 17.67 cm

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
                throw new FileNotFoundException($"Template '{_record.AwrNo}' not found in: {typeSubFolder}");

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

                // 1. UNPROTECT
                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    try { doc.Unprotect(WorkerConstants.RestrictEditPassword); } catch { }
                }

                // 2. RESIZE & CENTER
                SanitizeDocumentLayout(doc);

                // 3. STAMP
                foreach (Word.Section section in doc.Sections)
                {
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

                // 4. PROTECT & SAVE
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

        // --- RESIZING LOGIC ---
        private void SanitizeDocumentLayout(Word.Document doc)
        {
            Console.WriteLine(" > Adjusting Layout...");
            foreach (Word.Section section in doc.Sections)
            {
                // Set Margins & Distances (From Screenshot)
                section.PageSetup.TopMargin = PageMarginPt;
                section.PageSetup.BottomMargin = PageMarginPt;
                section.PageSetup.LeftMargin = PageMarginPt;
                section.PageSetup.RightMargin = PageMarginPt;
                section.PageSetup.HeaderDistance = HeaderDistPt;
                section.PageSetup.FooterDistance = FooterDistPt;

                // Process Body
                ResizeShapesInRange(section.Range);

                // Process Headers/Footers (Deep Search)
                foreach (Word.HeaderFooter hf in section.Headers) ResizeShapesInRange(hf.Range);
                foreach (Word.HeaderFooter hf in section.Footers) ResizeShapesInRange(hf.Range);
            }
        }

        private void ResizeShapesInRange(Word.Range range)
        {
            // A. Inline Shapes
            foreach (Word.InlineShape shape in range.InlineShapes)
            {
                // Force Unlock & Size
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.Height = TargetHeightPt;
                shape.Width = TargetWidthPt;
                shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            // B. Floating Shapes
            for (int i = range.ShapeRange.Count; i >= 1; i--)
            {
                var shape = range.ShapeRange[i];
                try
                {
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    shape.Height = TargetHeightPt;
                    shape.Width = TargetWidthPt;

                    // Absolute Center
                    shape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    shape.Top = (float)Word.WdShapePosition.wdShapeCenter;
                    shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    shape.Left = (float)Word.WdShapePosition.wdShapeCenter;

                    shape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBehindText);
                }
                catch { }
            }

            // C. Content Controls (Deep Search)
            foreach (Word.ContentControl cc in range.ContentControls)
            {
                ResizeShapesInRange(cc.Range);
            }
        }

        // ==========================================
        // 2. QC PRINTING
        // ==========================================
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
                // Narrow Margins for Receipt
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
                table.Columns[1].PreferredWidth = 30;
                table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[2].PreferredWidth = 70;
                table.Range.Font.Name = "Calibri";
                table.Range.Font.Size = 10;
                table.Range.ParagraphFormat.SpaceAfter = 3;
                table.Rows.WrapAroundText = 0;

                void AddRow(int r, string k, string v)
                {
                    table.Cell(r, 1).Range.Text = k;
                    table.Cell(r, 1).Range.Font.Bold = 1;
                    table.Cell(r, 2).Range.Text = v;
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
                default: return "";
            }
        }

        private string FindTemplateFile(string directory, string fileNameNoExt)
        {
            if (!Directory.Exists(directory)) return null;
            string pathDocx = Path.Combine(directory, fileNameNoExt + ".docx");
            if (File.Exists(pathDocx)) return pathDocx;
            string pathDoc = Path.Combine(directory, fileNameNoExt + ".doc");
            if (File.Exists(pathDoc)) return pathDoc;
            return null;
        }
    }
}