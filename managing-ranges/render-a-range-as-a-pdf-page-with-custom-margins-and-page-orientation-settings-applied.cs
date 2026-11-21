// Title: Render a selected Excel range to a single‑page PDF with custom margins using Aspose.Cells for .NET (C#)
// Common Searches: C# how to export only a specific Excel range to PDF with custom margins using Aspose.Cells | Aspose.Cells set print area and margins before saving worksheet as PDF in .NET | Generate a one‑page PDF from a defined cell range with Aspose.Cells C# example | Aspose.Cells PDFSaveOptions custom margins and single page per sheet tutorial | Export Excel A1:C5 range to PDF with custom page setup in C#
// Tags: Aspose.Cells range to PDF conversion | custom page margins Aspose.Cells | print area configuration Aspose.Cells | single page PDF Aspose.Cells | C# Aspose.Cells PDF export example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills cells A1:C5, defines that range as the print area, applies custom top, bottom, left, and right margins (in inches), configures PdfSaveOptions to keep the range on a single PDF page, and saves the result as 'RangeOutput.pdf' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in the range A1:C5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range that will be rendered to PDF (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range renderRange = sheet.Cells.CreateRange("A1:C5");

            // Set the print area to the defined range so only this area is exported
            sheet.PageSetup.PrintArea = renderRange.RefersTo;

            // Apply custom page margins (values are in inches)
            sheet.PageSetup.TopMargin = 0.5;      // 0.5 inch top margin
            sheet.PageSetup.BottomMargin = 0.5;   // 0.5 inch bottom margin
            sheet.PageSetup.LeftMargin = 0.75;    // 0.75 inch left margin
            sheet.PageSetup.RightMargin = 0.75;   // 0.75 inch right margin

            // Configure PDF save options to keep the range on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true   // Ensures the defined print area fits on one PDF page
            };

            // Define output file path
            string outputPath = "RangeOutput.pdf";

            // Save the worksheet (with the specified print area) as a PDF file
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF file saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
