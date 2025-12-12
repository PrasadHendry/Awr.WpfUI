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
        // CHANGED: Point to JPG which has Name embedded
        private readonly string _logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Sigma_Logo.jpg");

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

                // --- 1. SETUP ---
                ws.Protection.IsProtected = true; ws.Protection.SetPassword("QA");
                ws.Protection.AllowFormatColumns = true; ws.Protection.AllowFormatRows = true; ws.Protection.AllowSelectLockedCells = true;
                ws.DefaultRowHeight = 45;
                ws.Cells.Style.Font.Name = "Calibri"; ws.Cells.Style.Font.Size = 11; ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.View.FreezePanes(6, 1);

                // --- 2. HEADER CONTENT ---

                // Row 1: Image + Form No
                ws.Row(1).Height = 60;
                if (File.Exists(_logoPath))
                {
                    var logo = ws.Drawings.AddPicture("SigmaLogo", new FileInfo(_logoPath));
                    logo.SetSize(250, 55);
                    logo.SetPosition(0, 5, 0, 5);
                }

                ws.Cells["L1:M1"].Merge = true;
                ws.Cells["L1"].Value = "Form No.: SOP/QA/003/F2";
                ws.Cells["L1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["L1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                ws.Cells["L1"].Style.Font.Size = 9;

                // Row 2: Address
                ws.Cells["A2:M2"].Merge = true;
                ws.Cells["A2"].Value = "PLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA - 403526";
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                ws.Cells["A2"].Style.Indent = 1;
                ws.Cells["A2"].Style.Font.Size = 10;
                ws.Row(2).Height = 20;

                // Row 3: Title
                ws.Cells["A3:M3"].Merge = true;
                ws.Cells["A3"].Value = "Audit Trail for Issuance of AWR";
                ws.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A3"].Style.Font.Bold = true;
                ws.Cells["A3"].Style.Font.Size = 14;
                ws.Cells["A3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["A3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke);

                // Row 4: Group Headers
                ws.Cells["A4:F4"].Merge = true; ws.Cells["A4"].Value = "Request Details";
                ws.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; ws.Cells["A4"].Style.Font.Bold = true;

                ws.Cells["G4:M4"].Merge = true; ws.Cells["G4"].Value = "Issuance Details";
                ws.Cells["G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; ws.Cells["G4"].Style.Font.Bold = true;

                // --- 3. HEADER BORDERS (Explicit) ---

                // Main Outer Box (A1:M4)
                var headerBox = ws.Cells["A1:M4"];
                headerBox.Style.Border.BorderAround(ExcelBorderStyle.Thick);

                // Divider: Below Row 2 (Address)
                ws.Cells["A2:M2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // Divider: Below Row 3 (Title)
                ws.Cells["A3:M3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // Divider: Vertical between Groups (Row 4, Col F right side)
                ws.Cells["F4"].Style.Border.Right.Style = ExcelBorderStyle.Medium;

                // Row 5: Column Headers (Full Grid)
                for (int i = 0; i < _headers.Length; i++)
                {
                    var cell = ws.Cells[5, i + 1];
                    cell.Value = _headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.WrapText = true;
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                // Thicken Outer Edges of Row 5
                ws.Cells["A5"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                ws.Cells["M5"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                ws.Cells["A5:M5"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; // Separate headers from data

                // --- 4. DATA ROWS ---
                int row = 6;
                foreach (var item in data)
                {
                    // Values
                    ws.Cells[row, 1].Value = item.RequestNo;
                    ws.Cells[row, 2].Value = item.AwrNo;
                    ws.Cells[row, 3].Value = item.AwrType.ToString();
                    ws.Cells[row, 4].Value = item.MaterialProduct;
                    ws.Cells[row, 5].Value = item.BatchNo;
                    ws.Cells[row, 6].Value = item.ArNo;
                    ws.Cells[row, 7].Value = item.QtyIssued ?? item.QtyRequired;
                    ws.Cells[row, 8].Value = GetStatusDisplay(item.Status);
                    ws.Cells[row, 9].Value = FormatUserDate(item.RequestedBy, item.RequestedAt);
                    ws.Cells[row, 10].Value = FormatUserDate(item.IssuedBy, item.IssuedAt);
                    ws.Cells[row, 11].Value = FormatUserDate(item.ReceivedBy, item.ReceivedAt);
                    ws.Cells[row, 12].Value = FormatUserDate(item.ReturnedBy, item.ReturnedAt);
                    ws.Cells[row, 13].Value = item.Remark;

                    // Row Style
                    var rng = ws.Cells[row, 1, row, 13];
                    rng.Style.WrapText = true;
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                    // Inner Grid Borders
                    rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    // Thick Outer Borders
                    ws.Cells[row, 1].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                    ws.Cells[row, 13].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    // Alignment
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

                // Close Bottom Border
                ws.Cells[row - 1, 1, row - 1, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;

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
        // PDF EXPORT (QuestPDF - Updated Layout)
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

                    // HEADER (New Layout: Image Top, Text Below)
                    page.Header().Column(col =>
                    {
                        col.Item().Border(1.5f).BorderColor(Colors.Black).Padding(5).Row(row =>
                        {
                            // Left Section: Logo + Address Stack
                            row.RelativeItem().Column(c =>
                            {
                                // 1. Logo
                                if (File.Exists(_logoPath))
                                {
                                    byte[] logoBytes = File.ReadAllBytes(_logoPath);
                                    c.Item().Height(40).AlignLeft().Image(logoBytes).FitArea();
                                }

                                // 2. Address (Below Logo)
                                c.Item().Text("PLOT No. 6,7,8, TIVIM INDL. ESTATE, TIVIM, GOA - 403526").FontSize(8);
                            });

                            // Right Section: Form No
                            row.ConstantItem(150).AlignRight().AlignTop().Text("Form No.: SOP/QA/003/F2").FontSize(9);
                        });

                        col.Item().PaddingVertical(5);
                        col.Item().Background(Colors.Grey.Lighten4).Border(0.5f).BorderColor(Colors.Black).PaddingVertical(6).AlignCenter().Text("Audit Trail for Issuance of AWR").FontSize(14).Bold();
                        col.Item().PaddingVertical(5);
                    });

                    // TABLE (Same as before)
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