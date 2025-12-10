using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml; // EPPlus 4.5.3.3
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing; // For Excel Images
using Awr.Core.DTOs;
using Awr.Core.Enums;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace Awr.WpfUI.Services.Implementation
{
    public class ReportService
    {
        // Path to Logo (Must be Copy to Output Directory)
        private readonly string _logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "SIGMA_LOGO_BLK.png");

        public ReportService()
        {
            QuestPDF.Settings.License = LicenseType.Community;
        }

        // --- HELPERS ---

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

        private string FormatUserDate(string username, DateTime? date)
        {
            if (string.IsNullOrEmpty(username) || username == "NA") return "NA";
            if (!date.HasValue || date.Value.Year < 2000) return username;

            return $"{username}\n({date.Value:dd-MM-yyyy HH:mm})";
        }

        private readonly string[] _headers = {
            "Request No.", "AWR No.", "Type", "Material/Product", "Batch No.", "AR No.",
            "Qty Issued", "Status", "Requested By (QC)", "Approved By (QA)", "Printed By (QC)", "Voided By (QC)", "Remark / Justification"
        };

        // ==========================================
        // EXCEL EXPORT
        // ==========================================
        public void ExportToExcel(List<AwrItemQueueDto> data, string filePath, string username)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets.Add("Audit Trail");

                // Security
                ws.Protection.IsProtected = true; ws.Protection.SetPassword("QA");
                ws.Protection.AllowFormatColumns = true; ws.Protection.AllowFormatRows = true; ws.Protection.AllowSelectLockedCells = true;
                ws.DefaultRowHeight = 45;
                ws.Cells.Style.Font.Name = "Calibri"; ws.Cells.Style.Font.Size = 11; ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.View.FreezePanes(6, 1);

                // Header Frame
                ws.Cells["A1:M3"].Style.Border.BorderAround(ExcelBorderStyle.Thick);

                // Debug:
                if (!File.Exists(_logoPath))
                {
                    System.Windows.MessageBox.Show($"Logo missing at: {_logoPath}");
                }
                // --- INSERT LOGO ---
                if (File.Exists(_logoPath))
                {
                    var logo = ws.Drawings.AddPicture("SigmaLogo", new FileInfo(_logoPath));
                    logo.SetPosition(0, 5, 0, 5); // Row 0, Offset 5px
                    logo.SetSize(60, 60); // Adjust size as needed
                }


                // Company Text
                ws.Cells["A1:M1"].Merge = true;
                ws.Cells["A1"].Value = "SIGMA LABORATORIES PRIVATE LIMITED\nPLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA";
                ws.Cells["A1"].Style.WrapText = true; ws.Cells["A1"].Style.Font.Bold = true; ws.Cells["A1"].Style.Font.Size = 12;
                ws.Row(1).Height = 60;
                ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Form No
                ws.Cells["A2:M2"].Merge = true;
                ws.Cells["A2"].Value = "Form No.: SOP/QA/003/F2-02"; ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                // Title
                ws.Cells["A3:M3"].Merge = true;
                ws.Cells["A3"].Value = "Audit Trail for Issuance of AWR";
                ws.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; ws.Cells["A3"].Style.Font.Bold = true; ws.Cells["A3"].Style.Font.Size = 14;
                ws.Cells["A3"].Style.Fill.PatternType = ExcelFillStyle.Solid; ws.Cells["A3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke);

                // Group Headers
                ws.Cells["A4:F4"].Merge = true; ws.Cells["A4"].Value = "Request Details";
                ws.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; ws.Cells["A4"].Style.Border.BorderAround(ExcelBorderStyle.Thin); ws.Cells["A4"].Style.Font.Bold = true;

                ws.Cells["G4:M4"].Merge = true; ws.Cells["G4"].Value = "Issuance Details";
                ws.Cells["G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; ws.Cells["G4"].Style.Border.BorderAround(ExcelBorderStyle.Thin); ws.Cells["G4"].Style.Font.Bold = true;

                // Column Headers
                for (int i = 0; i < _headers.Length; i++)
                {
                    ws.Cells[5, i + 1].Value = _headers[i];
                    ws.Cells[5, i + 1].Style.Font.Bold = true;
                    ws.Cells[5, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    ws.Cells[5, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[5, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    ws.Cells[5, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[5, i + 1].Style.WrapText = true;
                }

                // Data Rows
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

                    // User + Date
                    ws.Cells[row, 9].Value = FormatUserDate(item.RequestedBy, item.RequestedAt);
                    ws.Cells[row, 10].Value = FormatUserDate(item.IssuedBy, item.IssuedAt);
                    ws.Cells[row, 11].Value = FormatUserDate(item.ReceivedBy, item.ReceivedAt);
                    ws.Cells[row, 12].Value = FormatUserDate(item.ReturnedBy, item.ReturnedAt);

                    ws.Cells[row, 13].Value = item.Remark;

                    // Styles
                    var rng = ws.Cells[row, 1, row, 13];
                    rng.Style.WrapText = true;
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    rng.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    // Alignments
                    ws.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[row, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    row++;
                }

                // Widths
                ws.Column(1).Width = 22; ws.Column(2).Width = 25; ws.Column(3).Width = 10;
                ws.Column(4).Width = 35; ws.Column(5).Width = 18; ws.Column(6).Width = 18;
                ws.Column(7).Width = 8; ws.Column(8).Width = 18; ws.Column(9).Width = 20;
                ws.Column(10).Width = 20; ws.Column(11).Width = 20; ws.Column(12).Width = 20;
                ws.Column(13).Width = 35;

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
                    page.DefaultTextStyle(x => x.FontSize(8).FontFamily("Calibri"));

                    // HEADER
                    page.Header().Column(col =>
                    {
                        col.Item().Border(1.5f).BorderColor(Colors.Black).Padding(0).Row(row =>
                        {
                            // 1. LOGO
                            if (File.Exists(_logoPath))
                            {
                                byte[] logoBytes = File.ReadAllBytes(_logoPath);
                                row.ConstantItem(60).Padding(5).AlignMiddle().Image(logoBytes);
                            }

                            // 2. Company Text
                            row.RelativeItem().Padding(10).Column(c =>
                            {
                                c.Item().Text("SIGMA LABORATORIES PRIVATE LIMITED").Bold().FontSize(11);
                                c.Item().Text("PLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA - 403526").FontSize(9);
                            });

                            // 3. Form No
                            row.ConstantItem(180).Padding(10).AlignRight().Text("Form No.: SOP/QA/003/F2-02").FontSize(10);
                        });

                        col.Item().PaddingVertical(5);
                        col.Item().Background(Colors.Grey.Lighten4).Border(0.5f).BorderColor(Colors.Black).PaddingVertical(6).AlignCenter().Text("Audit Trail for Issuance of AWR").FontSize(14).Bold();
                        col.Item().PaddingVertical(5);
                    });

                    // TABLE
                    page.Content().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.RelativeColumn(3); columns.RelativeColumn(3); columns.RelativeColumn(1);
                            columns.RelativeColumn(4); columns.RelativeColumn(2); columns.RelativeColumn(2);
                            columns.RelativeColumn(1); columns.RelativeColumn(2); columns.RelativeColumn(3);
                            columns.RelativeColumn(3); columns.RelativeColumn(3); columns.RelativeColumn(3); columns.RelativeColumn(4);
                        });

                        table.Header(header =>
                        {
                            foreach (var h in _headers) header.Cell().Element(HeaderCellStyle).Text(h).Bold();
                        });

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

                            // User + Date
                            table.Cell().Element(CenterDataStyle).Text(FormatUserDate(item.RequestedBy, item.RequestedAt));
                            table.Cell().Element(CenterDataStyle).Text(FormatUserDate(item.IssuedBy, item.IssuedAt));
                            table.Cell().Element(CenterDataStyle).Text(FormatUserDate(item.ReceivedBy, item.ReceivedAt));
                            table.Cell().Element(CenterDataStyle).Text(FormatUserDate(item.ReturnedBy, item.ReturnedAt));

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