using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Printing;       // [REQUIRED] Add Reference: System.Printing
using System.Windows.Forms;  // [REQUIRED] Add Reference: System.Windows.Forms
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

        // --- SPECIFICATIONS (UPDATED) ---
        // 1 cm = 28.35 points

        // 1. Page Margins (0.8 cm)
        private const float PageMarginPt = 22.68f;
        private const float HeaderDistPt = 22.68f;
        private const float FooterDistPt = 22.68f;

        // 2. Image Resizing (From New Screenshot)
        // Height: 24.98 cm -> 708.18 pt
        // Width:  18.99 cm -> 538.37 pt
        private const float TargetHeightPt = 708.18f;
        private const float TargetWidthPt = 538.37f;

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
            // 1. Locate Source File
            string typeSubFolder = GetSubFolderForType(_record.AwrType);
            string searchDirectory = Path.Combine(WorkerConstants.SourceRoot, typeSubFolder);
            string sourceFilePath = FindTemplateFile(searchDirectory, _record.AwrNo);

            if (string.IsNullOrEmpty(sourceFilePath))
                throw new FileNotFoundException($"Template '{_record.AwrNo}' not found in: {typeSubFolder}");

            // 2. Prepare Paths
            string extension = Path.GetExtension(sourceFilePath);
            string finalFileName = $"{_record.RequestNo}_{_record.AwrNo}{extension}";
            string finalFilePath = Path.Combine(WorkerConstants.FinalLocation, finalFileName);
            string tempFilePath = Path.Combine(WorkerConstants.TempLocation, Guid.NewGuid() + extension);

            // 3. Copy to Temp
            File.Copy(sourceFilePath, tempFilePath, true);

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);

                try { doc = wordApp.Documents.Open(tempFilePath, PasswordDocument: WorkerConstants.EncryptionPassword); }
                catch { doc = wordApp.Documents.Open(tempFilePath); }

                // 4. UNPROTECT
                if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    try { doc.Unprotect(WorkerConstants.RestrictEditPassword); } catch { }
                }

                // 5. RESIZE & CENTER (Uses updated dimensions 24.98cm x 18.99cm)
                SanitizeDocumentLayout(doc);

                // 6. APPLY STAMPS
                foreach (Word.Section section in doc.Sections)
                {
                    // =========================================================
                    // HEADER: Dark Blue Box "Stamp" (Fully Right Aligned)
                    // =========================================================
                    var headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Delete(); // Clear existing

                    // Create 1x1 Table for the Box
                    Word.Table stampTable = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                        .Range.Tables.Add(headerRange, 1, 1);

                    // FORCE RIGHT ALIGNMENT
                    stampTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowRight;

                    // Reset Indents to ensure it touches the margin
                    stampTable.Range.ParagraphFormat.RightIndent = 0;
                    stampTable.Range.ParagraphFormat.LeftIndent = 0;

                    // Fit the box tightly to text
                    stampTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

                    // Box Border Styling
                    stampTable.Borders.Enable = 1;
                    stampTable.Borders.OutsideColor = Word.WdColor.wdColorDarkBlue;
                    stampTable.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt; // Thick
                    stampTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    // Text Content & Styling
                    stampTable.Range.Text = _record.GetHeaderText();
                    stampTable.Range.Font.Name = "Arial";
                    stampTable.Range.Font.Size = 7; // Size 7
                    stampTable.Range.Font.Bold = 1;
                    stampTable.Range.Font.Color = Word.WdColor.wdColorDarkBlue;

                    // Center text *inside* the box
                    stampTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    stampTable.Range.ParagraphFormat.SpaceAfter = 0;

                    // =========================================================
                    // FOOTER: Digital Style with Top Line (Right Aligned)
                    // =========================================================
                    var footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Text = _record.GetFooterText();

                    // Text Styling
                    footerRange.Font.Name = "Consolas";
                    footerRange.Font.Size = 7; // Size 7
                    footerRange.Font.Color = Word.WdColor.wdColorBlack;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; // Right Align

                    // Gray Separator Line (Top)
                    footerRange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    footerRange.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                    footerRange.Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray50;

                    // Clear other borders
                    footerRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    footerRange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    footerRange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                }

                // 7. PROTECT & SAVE
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
                if (doc != null) { try { doc.Close(false); } catch { } Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { try { wordApp.Quit(); } catch { } Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
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
        // 2. QC PRINTING (SEQUENCE: DOC FIRST -> RECEIPT LAST)
        // ==========================================
        private void PrintSecureDocument()
        {
            string fileNameBase = $"{_record.RequestNo}_{_record.AwrNo}";
            string filePath = FindTemplateFile(WorkerConstants.FinalLocation, fileNameBase);

            if (string.IsNullOrEmpty(filePath)) throw new FileNotFoundException($"File not found: {fileNameBase}");

            // -----------------------------------------------------------------
            // STEP A: Prompt User for Printer
            // -----------------------------------------------------------------
            string selectedPrinterName = null;
            PrintDialog printDialog = new PrintDialog();
            printDialog.AllowSomePages = false;
            printDialog.AllowSelection = false;

            // Note: Ensure [STAThread] is in Program.cs for this to show!
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                selectedPrinterName = printDialog.PrinterSettings.PrinterName;
            }
            else
            {
                throw new Exception("Printing Cancelled by User.");
            }

            // -----------------------------------------------------------------
            // STEP B: Force Printer Settings (Simplex + A4)
            // -----------------------------------------------------------------
            ForcePrinterSettings(selectedPrinterName);

            // -----------------------------------------------------------------
            // STEP C: Word Automation
            // -----------------------------------------------------------------
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                Program.ActiveWordApps.Add(wordApp);

                // Explicitly set the printer
                wordApp.ActivePrinter = selectedPrinterName;

                // -------------------------------------------------------------
                // 1. PRINT MAIN DOCUMENT (FIRST)
                // -------------------------------------------------------------
                Console.WriteLine($" > Printing Main Doc ({_record.QtyIssued:0} Copies) to {selectedPrinterName}...");

                doc = wordApp.Documents.Open(filePath, PasswordDocument: WorkerConstants.EncryptionPassword, ReadOnly: true);

                int copies = (int)_record.QtyIssued;
                if (copies < 1) copies = 1;

                // The Printer Driver now has the "OneSided" ticket forced from Step B
                doc.PrintOut(Background: false, Copies: copies);

                // CLOSE the main document immediately after sending to spooler
                doc.Close(false);
                Marshal.ReleaseComObject(doc);
                doc = null;

                // -------------------------------------------------------------
                // 2. PRINT RECEIPT (LAST)
                // -------------------------------------------------------------
                // Now that the main doc is closed, we use the same Word App to print the receipt
                PrintReceiptTable(wordApp, selectedPrinterName);
            }
            finally
            {
                if (doc != null) { try { doc.Close(false); } catch { } Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); Program.ActiveWordApps.Remove(wordApp); }
            }
        }

        /// <summary>
        /// Modifies the Print Ticket of the selected printer to enforce A4 and One-Sided printing.
        /// </summary>
        private void ForcePrinterSettings(string printerName)
        {
            try
            {
                using (LocalPrintServer printServer = new LocalPrintServer())
                {
                    using (PrintQueue queue = printServer.GetPrintQueue(printerName))
                    {
                        // 1. Get the User's current Ticket
                        PrintTicket userTicket = queue.UserPrintTicket;

                        // 2. Force One Sided (Simplex)
                        if (userTicket.Duplexing.HasValue)
                        {
                            userTicket.Duplexing = Duplexing.OneSided;
                            Console.WriteLine(" > Enforced: One-Sided Print");
                        }

                        // 3. Force A4
                        if (userTicket.PageMediaSize != null)
                        {
                            userTicket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
                            Console.WriteLine(" > Enforced: A4 Paper Size");
                        }

                        // 4. Commit changes to the queue for this session
                        queue.UserPrintTicket = userTicket;
                        queue.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($" ! Warning: Could not enforce printer settings (Driver restrictions?): {ex.Message}");
                // We do not throw here to allow printing to continue even if enforcing fails.
            }
        }

        private void PrintReceiptTable(Word.Application wordApp, string printerName)
        {
            Console.WriteLine(" > Printing Receipt...");
            Word.Document doc = wordApp.Documents.Add();

            try
            {
                // Ensure Receipt uses the same printer
                wordApp.ActivePrinter = printerName;

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