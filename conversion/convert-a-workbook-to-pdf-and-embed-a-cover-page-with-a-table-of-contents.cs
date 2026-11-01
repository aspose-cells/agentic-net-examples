// Title: Generate a PDF from an Excel workbook with a custom cover sheet and hyperlinked table of contents using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET add a cover worksheet and TOC before saving Excel as PDF | C# create PDF from Excel with clickable table of contents using Aspose.Cells | How to insert a title page and sheet hyperlinks when converting XLSX to PDF in Aspose.Cells | Generate PDF with a title page and navigation links from an Excel file in C# | Aspose.Cells PDFSaveOptions example with cover page and table of contents
// Tags: Aspose.Cells add cover worksheet for PDF export | Excel to PDF with clickable table of contents using C# | PdfSaveOptions OnePagePerSheet false Aspose.Cells | hyperlink sheet names in PDF generated from Excel | auto-fit columns on cover sheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing XLSX file, inserts a new worksheet as a cover page, writes a report title and a table of contents that lists each subsequent sheet with hyperlinks to their A1 cells, auto‑fits the cover columns, and saves the workbook as a PDF using PdfSaveOptions with OnePagePerSheet disabled and PDF/A‑1b compliance.
class WorkbookToPdfWithCover
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Insert a new worksheet at the beginning to serve as the cover page
            Worksheet coverSheet = workbook.Worksheets.Insert(0, SheetType.Worksheet);
            coverSheet.Name = "Cover";

            // ----- Build Cover Page -----
            // Title
            Cell titleCell = coverSheet.Cells["A1"];
            titleCell.PutValue("Report Title");
            Style titleStyle = titleCell.GetStyle();
            titleStyle.Font.Size = 24;
            titleStyle.Font.IsBold = true;
            titleCell.SetStyle(titleStyle);

            // Table of Contents header
            Cell tocHeader = coverSheet.Cells["A3"];
            tocHeader.PutValue("Table of Contents");
            Style headerStyle = tocHeader.GetStyle();
            headerStyle.Font.Size = 16;
            headerStyle.Font.IsBold = true;
            tocHeader.SetStyle(headerStyle);

            // ----- Populate Table of Contents -----
            int tocStartRow = 4; // Row where TOC entries begin
            for (int i = 1; i < workbook.Worksheets.Count; i++) // start from 1 to skip the cover sheet
            {
                Worksheet sheet = workbook.Worksheets[i];
                string sheetName = sheet.Name;

                // Write sheet name in the TOC
                Cell tocEntry = coverSheet.Cells[tocStartRow + i - 1, 0]; // Column A
                tocEntry.PutValue(sheetName);

                // Add hyperlink to the first cell of the target sheet
                // The hyperlink format: 'SheetName'!A1
                string hyperlinkAddress = $"'{sheetName}'!A1";
                coverSheet.Hyperlinks.Add(tocEntry.Row, tocEntry.Column, 1, 1, hyperlinkAddress);
            }

            // Optional: Adjust column width for better appearance
            coverSheet.AutoFitColumns();

            // ----- Save as PDF -----
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,               // Allow multiple pages per sheet if needed
                Compliance = PdfCompliance.PdfA1b      // Example compliance setting
            };

            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
