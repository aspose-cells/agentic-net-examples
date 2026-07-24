// Title: Add a Custom Logo Cover Page and Convert an Aspose.Cells Workbook to PDF (C#)
// Description: Creates a workbook, inserts a centered logo, title and subtitle on a cover worksheet, then saves all visible sheets to a PDF using Aspose.Cells PdfSaveOptions.
// Keywords: Aspose.Cells | C# | PDF conversion | cover page | logo insertion | workbook to PDF | PdfSaveOptions | SheetSet.Visible | Excel to PDF | .NET | image positioning | page setup
// Common Searches: Aspose.Cells add cover page with logo before PDF export | C# convert Excel workbook to PDF with custom title page | How to insert and center an image in an Aspose.Cells worksheet | PdfSaveOptions include only visible sheets Aspose.Cells | Create PDF report with branding using Aspose.Cells C#
// Developer Intent: Generate a PDF that starts with a branded cover sheet containing a logo, title and subtitle, followed by the workbook’s data sheets.
// Use Cases: Corporate annual reports that need a branded first page | Automated PDF brochures with a designed cover before data tables | Standardized PDF exports from multiple workbooks with consistent branding
// AI Prompts: Show C# code to add a background image to the cover sheet and adjust margins before PDF conversion with Aspose.Cells. | Explain how to hide gridlines on the cover page while keeping them visible on data sheets using PdfSaveOptions. | Provide an example of loading a logo from a URL at runtime and embedding it in the cover worksheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsCoverPdfDemo
{
    // Creates a workbook, inserts a centered logo, title and subtitle on a cover worksheet, then saves all visible sheets to a PDF using Aspose.Cells PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the custom logo image (ensure the file exists)
                string logoPath = "logo.png";

                // Output PDF file path
                string outputPdf = "WorkbookWithCover.pdf";

                // -------------------- Create a new workbook --------------------
                Workbook workbook = new Workbook();

                // -------------------- Add cover page (first worksheet) --------------------
                Worksheet coverSheet = workbook.Worksheets[0];
                coverSheet.Name = "Cover";

                // Insert the logo picture if the file exists
                if (File.Exists(logoPath))
                {
                    int pictureIndex = coverSheet.Pictures.Add(0, 0, logoPath);
                    Picture logo = coverSheet.Pictures[pictureIndex];
                    // Set picture size (in points)
                    logo.Width = 200;
                    logo.Height = 100;
                    // Approximate centering of the picture
                    logo.Left = (coverSheet.Cells.MaxColumn + 1) * 64 / 2 - logo.Width / 2;
                }
                else
                {
                    Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Skipping logo insertion.");
                }

                // Add title text below the logo
                Cell titleCell = coverSheet.Cells["A5"];
                titleCell.PutValue("Annual Report 2026");
                Style titleStyle = workbook.CreateStyle();
                titleStyle.Font.Size = 24;
                titleStyle.Font.IsBold = true;
                titleStyle.HorizontalAlignment = TextAlignmentType.Center;
                titleCell.SetStyle(titleStyle);
                // Merge cells for the title to span across columns (A5:F5)
                coverSheet.Cells.Merge(4, 0, 1, 5);

                // Optionally add subtitle or date
                Cell subtitleCell = coverSheet.Cells["A7"];
                subtitleCell.PutValue($"Generated on {DateTime.Now:MMMM dd, yyyy}");
                Style subStyle = workbook.CreateStyle();
                subStyle.Font.Size = 12;
                subStyle.HorizontalAlignment = TextAlignmentType.Center;
                subtitleCell.SetStyle(subStyle);
                // Merge cells for the subtitle (A7:F7)
                coverSheet.Cells.Merge(6, 0, 1, 5);

                // Set page setup for cover sheet (fit to one page)
                coverSheet.PageSetup.FitToPagesWide = 1;
                coverSheet.PageSetup.FitToPagesTall = 1;

                // -------------------- Add main content sheet --------------------
                int dataSheetIndex = workbook.Worksheets.Add();
                Worksheet dataSheet = workbook.Worksheets[dataSheetIndex];
                dataSheet.Name = "Data";

                // Populate sample data
                dataSheet.Cells["A1"].PutValue("Item");
                dataSheet.Cells["B1"].PutValue("Quantity");
                dataSheet.Cells["A2"].PutValue("Apples");
                dataSheet.Cells["B2"].PutValue(150);
                dataSheet.Cells["A3"].PutValue("Bananas");
                dataSheet.Cells["B3"].PutValue(200);
                dataSheet.Cells["A4"].PutValue("Cherries");
                dataSheet.Cells["B4"].PutValue(75);

                // Apply simple table style to header row
                Style tableStyle = workbook.CreateStyle();
                tableStyle.Font.IsBold = true;
                dataSheet.Cells["A1:B1"].SetStyle(tableStyle);
                dataSheet.AutoFitColumns();

                // -------------------- Save workbook as PDF with options --------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Save all visible sheets; cover sheet is first
                    SheetSet = SheetSet.Visible,
                    // Allow normal pagination for data sheet
                    OnePagePerSheet = false
                };

                // Save to PDF
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"Workbook successfully saved to PDF with cover page: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
