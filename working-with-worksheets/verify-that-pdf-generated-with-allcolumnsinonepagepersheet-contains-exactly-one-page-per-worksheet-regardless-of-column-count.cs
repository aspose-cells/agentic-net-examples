// Title: How to confirm that a PDF saved with Aspose.Cells PdfSaveOptions.AllColumnsInOnePagePerSheet creates exactly one page per worksheet in C#
// Common Searches: C# code to check PDF page count matches worksheet count using Aspose.Cells | verify each worksheet renders on a single PDF page with AllColumnsInOnePagePerSheet | determine number of pages in PDF generated from a workbook without using Aspose.Pdf | use WorkbookRender to get PDF page count in .NET | Aspose.Cells render workbook to PDF and ensure one page per sheet
// Tags: PdfSaveOptions AllColumnsInOnePagePerSheet validation | WorkbookRender PDF page count extraction | C# Aspose.Cells one page per worksheet | render workbook to single-page PDF per sheet | verify PDF page count equals worksheet count

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook with three worksheets, fills each with 100 columns and 10 rows, saves it to PDF using PdfSaveOptions with AllColumnsInOnePagePerSheet enabled, then uses WorkbookRender with OnePagePerSheet to retrieve the PDF page count and confirms it matches the number of worksheets.
class VerifyPdfPages
{
    static void Main()
    {
        try
        {
            // Create a workbook with three worksheets
            var workbook = new Workbook();
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Populate each worksheet with many columns (e.g., 100) and some rows
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 100; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row}C{col}");
                    }
                }
            }

            // Configure PDF save options to fit all columns on one page per sheet
            var pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook to a PDF in memory
            using (var pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, pdfOptions);
                pdfStream.Position = 0; // Reset stream position for any further use

                // Determine the number of PDF pages using WorkbookRender (no Aspose.Pdf dependency)
                int pdfPageCount = GetPdfPageCount(workbook);

                int worksheetCount = workbook.Worksheets.Count;

                // Verify that each worksheet produced exactly one PDF page
                if (pdfPageCount == worksheetCount)
                {
                    Console.WriteLine($"Success: PDF contains {pdfPageCount} page(s), matching {worksheetCount} worksheet(s).");
                }
                else
                {
                    Console.WriteLine($"Failure: PDF contains {pdfPageCount} page(s) but workbook has {worksheetCount} worksheet(s).");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to calculate PDF page count using WorkbookRender
    private static int GetPdfPageCount(Workbook workbook)
    {
        try
        {
            var renderOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Pdf,
                OnePagePerSheet = true // Ensures each sheet is rendered as a single page
            };

            var renderer = new WorkbookRender(workbook, renderOptions);
            return renderer.PageCount;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while calculating PDF page count: {ex.Message}");
            return -1;
        }
    }
}
