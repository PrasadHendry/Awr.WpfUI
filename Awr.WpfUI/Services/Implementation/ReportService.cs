using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml; // EPPlus 4.5.3.3
using OfficeOpenXml.Style;
using Awr.Core.DTOs;
using Awr.Core.Enums;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace Awr.WpfUI.Services.Implementation
{
    public class ReportService
    {
        public ReportService()
        {
            QuestPDF.Settings.License = LicenseType.Community;
        }

        private string GetStatusDisplay(AwrItemStatus s)
        {
            switch (s)
            {
                case AwrItemStatus.PendingIssuance: return "Pending Approval";
                case AwrItemStatus.Issued: return "Approved";
                case AwrItemStatus.Received: return "Completed";
                case AwrItemStatus.Voided: return "Voided";
                case AwrItemStatus.RejectedByQa: return "Rejected";
                default: return s.ToString();
            }
        }

        // ==========================================
        // EXCEL EXPORT
        // ==========================================
        public void ExportToExcel(List<AwrItemQueueDto> data, string filePath, string username)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets.Add("Audit Trail");

                // --- 1. SECURITY (Protected but User-Friendly) ---
                ws.Protection.IsProtected = true;
                ws.Protection.SetPassword("QA");
                ws.Protection.AllowFormatColumns = true;
                ws.Protection.AllowFormatRows = true;
                ws.Protection.AllowSelectLockedCells = true;

                // --- 2. GLOBAL STYLING ---
                ws.DefaultRowHeight = 35; // Comfortable height
                ws.Cells.Style.Font.Name = "Calibri";
                ws.Cells.Style.Font.Size = 11;
                ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.View.FreezePanes(6, 1); // Freeze Header Rows

                // --- 3. HEADER SECTION ---
                // Box Border around Header
                ws.Cells["A1:L3"].Style.Border.BorderAround(ExcelBorderStyle.Thick);

                // Company Name
                ws.Cells["A1:L1"].Merge = true;
                ws.Cells["A1"].Value = "SIGMA LABORATORIES PRIVATE LIMITED\nPLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA";
                ws.Cells["A1"].Style.WrapText = true;
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["A1"].Style.Font.Size = 12;
                ws.Row(1).Height = 60;

                // Form No
                ws.Cells["A2:L2"].Merge = true;
                ws.Cells["A2"].Value = "Form No.: SOP/QA/003/F2-02";
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                // Title
                ws.Cells["A3:L3"].Merge = true;
                ws.Cells["A3"].Value = "Audit Trail for Issuance of AWR";
                ws.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A3"].Style.Font.Bold = true;
                ws.Cells["A3"].Style.Font.Size = 14;
                ws.Cells["A3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["A3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke);

                // Group Headers
                ws.Cells["A4:F4"].Merge = true; ws.Cells["A4"].Value = "Request Details";
                ws.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                ws.Cells["G4:L4"].Merge = true; ws.Cells["G4"].Value = "Issuance Details";
                ws.Cells["G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["G4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                // Column Headers
                string[] headers = { "Request No", "AWR No", "Type", "Material", "Batch", "AR No", "Qty", "Status", "Prepared By", "Approved By", "Printed By", "Remark" };
                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cells[5, i + 1].Value = headers[i];
                    ws.Cells[5, i + 1].Style.Font.Bold = true;
                    ws.Cells[5, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    ws.Cells[5, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[5, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    ws.Cells[5, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                // --- 4. DATA ROWS ---
                int row = 6;
                foreach (var item in data)
                {
                    ws.Cells[row, 1].Value = item.RequestNo;
                    ws.Cells[row, 2].Value = item.AwrNo;
                    ws.Cells[row, 3].Value = item.AwrType.ToString();
                    ws.Cells[row, 4].Value = item.MaterialProduct;
                    ws.Cells[row, 5].Value = item.BatchNo;
                    ws.Cells[row, 6].Value = item.ArNo;
                    ws.Cells[row, 7].Value = item.QtyIssued ?? item.QtyRequired;
                    ws.Cells[row, 8].Value = GetStatusDisplay(item.Status);
                    ws.Cells[row, 9].Value = item.RequestedBy;
                    ws.Cells[row, 10].Value = item.IssuedBy;
                    ws.Cells[row, 11].Value = item.ReceivedBy;
                    ws.Cells[row, 12].Value = item.Remark;

                    // Common Row Style
                    var rng = ws.Cells[row, 1, row, 12];
                    rng.Style.WrapText = true;
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Top; // Top align for multi-line content
                    rng.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    // SPECIFIC ALIGNMENTS
                    // Center: Req, Type, Qty, Status, Users
                    ws.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // Left: AWR, Mat, Batch, AR, Remark
                    ws.Cells[row, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    row++;
                }

                // --- 5. WIDTHS ---
                ws.Column(1).Width = 22; // Req
                ws.Column(2).Width = 25; // AWR
                ws.Column(3).Width = 10; // Type
                ws.Column(4).Width = 40; // Material (Wide)
                ws.Column(5).Width = 20; // Batch
                ws.Column(6).Width = 20; // AR
                ws.Column(7).Width = 8;  // Qty
                ws.Column(8).Width = 18; // Status
                ws.Column(9).Width = 22; // Prep
                ws.Column(10).Width = 22; // Appr
                ws.Column(11).Width = 22; // Prnt
                ws.Column(12).Width = 35; // Remark

                // Footer
                row += 2;
                ws.Cells[row, 1].Value = $"Generated By: {username} | Generated On: {DateTime.Now:dd-MM-yyyy HH:mm}";
                ws.Cells[row, 1].Style.Font.Italic = true;

                package.Save();
            }
        }

        // ==========================================
        // PDF EXPORT (QuestPDF)
        // ==========================================

        private IContainer HeaderCellStyle(IContainer c) => c.Background(Colors.Grey.Lighten3).Border(0.5f).BorderColor(Colors.Black).Padding(4).AlignCenter().AlignMiddle();

        private IContainer LeftDataStyle(IContainer c) => c.Border(0.5f).BorderColor(Colors.Grey.Medium).Padding(4).AlignLeft().AlignTop();

        private IContainer CenterDataStyle(IContainer c) => c.Border(0.5f).BorderColor(Colors.Grey.Medium).Padding(4).AlignCenter().AlignTop();

        public void ExportToPdf(List<AwrItemQueueDto> data, string filePath, string username)
        {
            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(9).FontFamily("Calibri"));

                    // HEADER
                    page.Header().Column(col =>
                    {
                        col.Item().Border(1.5f).BorderColor(Colors.Black).Padding(0).Row(row =>
                        {
                            // Company
                            row.RelativeItem().Padding(10).Column(c =>
                            {
                                c.Item().Text("SIGMA LABORATORIES PRIVATE LIMITED").Bold().FontSize(12);
                                c.Item().Text("PLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA - 403526").FontSize(10);
                            });
                            // Form No
                            row.ConstantItem(180).Padding(10).AlignRight().Text("Form No.: SOP/QA/003/F2-02").FontSize(10);
                        });

                        col.Item().PaddingVertical(5);
                        // Title Bar
                        col.Item().Background(Colors.Grey.Lighten4).Border(0.5f).BorderColor(Colors.Black).PaddingVertical(6).AlignCenter().Text("Audit Trail for Issuance of AWR").FontSize(14).Bold();
                        col.Item().PaddingVertical(5);
                    });

                    // TABLE
                    page.Content().Table(table =>
                    {
                        // Column Definitions (Total 12)
                        table.ColumnsDefinition(columns =>
                        {
                            columns.RelativeColumn(3); // Req
                            columns.RelativeColumn(3); // AWR
                            columns.RelativeColumn(1); // Type
                            columns.RelativeColumn(4); // Mat
                            columns.RelativeColumn(2); // Batch
                            columns.RelativeColumn(2); // AR
                            columns.RelativeColumn(1); // Qty
                            columns.RelativeColumn(2); // Status
                            columns.RelativeColumn(3); // Prep
                            columns.RelativeColumn(3); // Appr
                            columns.RelativeColumn(3); // Prnt
                            columns.RelativeColumn(4); // Rem
                        });

                        // Header Row
                        table.Header(header =>
                        {
                            header.Cell().Element(HeaderCellStyle).Text("Request No").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("AWR No").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Type").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Material").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Batch").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("AR No").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Qty").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Status").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Prepared By").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Approved By").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Printed By").Bold();
                            header.Cell().Element(HeaderCellStyle).Text("Remark").Bold();
                        });

                        // Data Rows
                        foreach (var item in data)
                        {
                            table.Cell().Element(CenterDataStyle).Text(item.RequestNo);
                            table.Cell().Element(LeftDataStyle).Text(item.AwrNo);
                            table.Cell().Element(CenterDataStyle).Text(item.AwrType.ToString());
                            table.Cell().Element(LeftDataStyle).Text(item.MaterialProduct);
                            table.Cell().Element(LeftDataStyle).Text(item.BatchNo);
                            table.Cell().Element(LeftDataStyle).Text(item.ArNo);
                            table.Cell().Element(CenterDataStyle).Text(item.QtyIssued?.ToString("0") ?? "0");
                            table.Cell().Element(CenterDataStyle).Text(GetStatusDisplay(item.Status));
                            table.Cell().Element(CenterDataStyle).Text(item.RequestedBy);
                            table.Cell().Element(CenterDataStyle).Text(item.IssuedBy);
                            table.Cell().Element(CenterDataStyle).Text(item.ReceivedBy);
                            table.Cell().Element(LeftDataStyle).Text(item.Remark);
                        }
                    });

                    // FOOTER
                    page.Footer().PaddingTop(10).Row(row =>
                    {
                        row.RelativeItem().Text(x => { x.Span("Generated By: "); x.Span(username).Italic(); x.Span($" | {DateTime.Now}"); });
                        row.RelativeItem().AlignRight().Text(x => { x.CurrentPageNumber(); x.Span(" / "); x.TotalPages(); });
                    });
                });
            })
            .GeneratePdf(filePath);
        }
    }
}