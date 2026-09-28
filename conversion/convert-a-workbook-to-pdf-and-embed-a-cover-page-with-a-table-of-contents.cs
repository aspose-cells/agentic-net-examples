// Title: Create a cover worksheet with a table of contents and export an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts a cover sheet, builds a simple table of contents listing each worksheet, and saves the workbook as a PDF with Aspose.Cells. | Demonstrate how to configure PdfSaveOptions to export an Excel file to PDF after adding a formatted cover page and applying AutoFit to columns.
// Common Searches: how to add a cover page to an Excel file before converting to PDF with Aspose.Cells C# | Aspose.Cells generate table of contents worksheet for PDF export | C# export workbook to PDF with custom cover sheet using PdfSaveOptions | placeholder page numbers in TOC worksheet Aspose.Cells
// Tags: cover worksheet insertion Aspose.Cells | table of contents generation Excel PDF | PdfSaveOptions configuration Aspose.Cells | auto fit columns before PDF conversion | placeholder page numbers TOC Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example checks for the source .xlsx file, inserts a new worksheet at the beginning as a cover page, writes a title and a table of contents with placeholder page numbers for each subsequent sheet, auto‑fits columns, sets PdfSaveOptions, and saves the complete workbook as a PDF.
class WorkbookToPdfWithCover
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputDocument.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Insert a new worksheet at the beginning to serve as the cover page
            Worksheet coverSheet = workbook.Worksheets.Insert(0, SheetType.Worksheet);
            coverSheet.Name = "Cover";

            // Prepare the cover page content
            Cells cells = coverSheet.Cells;

            // Title
            cells["A1"].PutValue("Document Title");
            Style titleStyle = workbook.CreateStyle();
            titleStyle.Font.IsBold = true;
            titleStyle.Font.Size = 20;
            cells["A1"].SetStyle(titleStyle);

            // Table of Contents header
            cells["A3"].PutValue("Table of Contents");
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 14;
            cells["A3"].SetStyle(headerStyle);

            // List each worksheet (excluding the cover) as an entry in the TOC
            int row = 4;
            for (int i = 1; i < workbook.Worksheets.Count; i++) // start from 1 to skip cover
            {
                Worksheet ws = workbook.Worksheets[i];
                // Entry format: Sheet Name .......... Page #
                // Page numbers are not known before rendering; placeholder used.
                cells[row, 0].PutValue(ws.Name);
                cells[row, 1].PutValue("Page ..."); // placeholder
                row++;
            }

            // Adjust column widths for better appearance
            coverSheet.AutoFitColumns();

            // Set PDF save options (optional: one page per sheet, etc.)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false // allow sheets to span multiple pages if needed
            };

            // Save the workbook (including the cover page) as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
