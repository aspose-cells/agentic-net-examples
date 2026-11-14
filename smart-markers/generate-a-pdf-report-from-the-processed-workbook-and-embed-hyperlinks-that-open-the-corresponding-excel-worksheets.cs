// Title: Create a PDF report from a processed Excel workbook with internal worksheet hyperlinks using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add hyperlink to another worksheet before saving as PDF | How to embed internal worksheet links in a PDF generated from Excel with Aspose.Cells | Save Excel workbook as PDF/A-1b with navigation links using Aspose.Cells in .NET | Aspose.Cells PdfSaveOptions fit columns to page width while keeping sheet layout | Generate PDF report from a processed workbook and link to a specific cell using Aspose.Cells
// Tags: add internal worksheet hyperlink Aspose.Cells | PdfSaveOptions column fit per sheet | save workbook as PDF/A-1b Aspose.Cells | generate PDF report from Excel Aspose.Cells | embed worksheet navigation links PDF Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The example loads a processed Excel workbook, verifies it contains at least two worksheets, adds an internal hyperlink from cell A1 on the first sheet to cell B2 on the second sheet, sets the display text, configures PdfSaveOptions to keep the original sheet layout, fit all columns to the page width, and optionally enforce PDF/A-1b compliance, then saves the workbook as a PDF report with the embedded navigation hyperlink.
class PdfReportGenerator
{
    static void Main()
    {
        try
        {
            const string inputFile = "ProcessedWorkbook.xlsx";
            const string outputFile = "Report.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the processed workbook
            Workbook workbook = new Workbook(inputFile);

            // Ensure there are at least two worksheets for the example
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("Error: The workbook must contain at least two worksheets.");
                return;
            }

            // Reference worksheets
            Worksheet firstSheet = workbook.Worksheets[0];
            Worksheet targetSheet = workbook.Worksheets[1];

            // Add an internal hyperlink pointing to the target sheet cell B2
            // Use the 5‑parameter overload (row, column, rowsCount, columnsCount, address)
            firstSheet.Hyperlinks.Add(0, 0, 1, 1, $"{targetSheet.Name}!B2");

            // Set display text for the hyperlink
            firstSheet.Cells["A1"].PutValue($"Go to {targetSheet.Name} B2");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,               // Keep original sheet layout
                AllColumnsInOnePagePerSheet = true,    // Fit columns to page width
                Compliance = PdfCompliance.PdfA1b      // PDF/A-1b compliance (optional)
            };

            // Save the workbook as a PDF report with embedded hyperlinks
            workbook.Save(outputFile, pdfOptions);
            Console.WriteLine($"PDF report generated successfully: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
