// Title: Add a Cover Sheet with Logo and Title, then Export Workbook to PDF/A‑1b using Aspose.Cells for .NET (C#)
// Description: Shows how to create or load an Excel workbook, insert a cover worksheet as the first sheet, embed a PNG logo, add a centered bold title, adjust column width and page setup, add a data sheet, configure PdfSaveOptions for PDF/A‑1b compliance and default‑font checking, and save the file as a PDF where the cover page appears first.
// Keywords: Aspose.Cells C# PDF conversion | cover page Excel Aspose | embed image Excel worksheet | PDF/A-1b Aspose.Cells | PdfSaveOptions OnePagePerSheet | add logo to Excel | export workbook to PDF | custom cover sheet PDF | Aspose.Cells page setup | C# Excel to PDF with logo
// Common Searches: Aspose.Cells add cover page before PDF export | C# embed logo in Excel and save as PDF/A | How to create PDF/A‑1b from Excel with Aspose.Cells | Set first worksheet as cover sheet in Aspose.Cells | Configure PdfSaveOptions for PDF/A compliance in .NET | Insert image into Excel cell using Aspose.Cells C#
// Developer Intent: Create a workbook with a branded cover sheet and export it as a PDF/A‑1b document.
// Use Cases: Branding corporate reports with a logo‑filled cover page | Generating compliant PDF/A‑1b archives of financial spreadsheets | Automating multi‑sheet PDF creation where the first page is a custom title page | Producing printable reports that require a separate cover sheet layout
// AI Prompts: Generate C# code that uses Aspose.Cells to insert a PNG logo into the first worksheet, add a centered bold title, and save the workbook as a PDF/A‑1b file. | Show how to set PdfSaveOptions such as Compliance = PdfA1b, CheckWorkbookDefaultFont = true, and OnePagePerSheet = false when converting Excel to PDF with Aspose.Cells. | Provide a complete example that creates a cover sheet, adds sample data, handles a missing logo file, and exports the workbook to PDF with proper page setup.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Shows how to create or load an Excel workbook, insert a cover worksheet as the first sheet, embed a PNG logo, add a centered bold title, adjust column width and page setup, add a data sheet, configure PdfSaveOptions for PDF/A‑1b compliance and default‑font checking, and save the file as a PDF where the cover page appears first.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // 2. Add a cover sheet as the first worksheet
            // Insert a new empty worksheet at index 0 (first position)
            workbook.Worksheets.Insert(0, SheetType.Worksheet); // Correct overload
            Worksheet coverSheet = workbook.Worksheets[0];
            coverSheet.Name = "Cover";

            // 3. Insert a custom logo image into the cover sheet
            string logoPath = "logo.png";
            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Logo file not found: {logoPath}");
                return;
            }

            // Load the image into a stream and add it to cell A1
            using (FileStream logoStream = File.OpenRead(logoPath))
            {
                coverSheet.Pictures.Add(0, 0, logoStream);
            }

            // 4. Add title text (or any other cover page content)
            Cell titleCell = coverSheet.Cells["A5"];
            titleCell.PutValue("Annual Report");

            // Apply style to the title cell
            Style titleStyle = titleCell.GetStyle();
            titleStyle.Font.IsBold = true;
            titleStyle.Font.Size = 24;
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            titleCell.SetStyle(titleStyle);

            // Widen column A for better appearance
            coverSheet.Cells.SetColumnWidth(0, 30);

            // 5. (Optional) Configure page setup for the cover sheet
            coverSheet.PageSetup.FitToPagesWide = 1;
            coverSheet.PageSetup.FitToPagesTall = 1;

            // 6. Add a regular data worksheet (example content)
            Worksheet dataSheet = workbook.Worksheets[1];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].PutValue("Item");
            dataSheet.Cells["B1"].PutValue("Quantity");
            dataSheet.Cells["A2"].PutValue("Apples");
            dataSheet.Cells["B2"].PutValue(150);
            dataSheet.Cells["A3"].PutValue("Oranges");
            dataSheet.Cells["B3"].PutValue(200);

            // 7. Set PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep each worksheet on its own page(s)
                OnePagePerSheet = false,

                // Ensure default font handling for Unicode characters
                CheckWorkbookDefaultFont = true,

                // Produce PDF/A‑1b compliant file
                Compliance = PdfCompliance.PdfA1b
            };

            // 8. Save the workbook as a PDF file; the cover sheet will be the first page(s)
            string outputPdf = "ReportWithCover.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPdf}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
