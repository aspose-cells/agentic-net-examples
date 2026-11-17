// Title: Generate PDF bookmarks from Excel worksheet names using Aspose.Cells PdfSaveOptions in C#
// Common Searches: how to enable worksheet bookmarks when saving Excel as PDF with Aspose.Cells C# | Aspose.Cells PdfSaveOptions Bookmarks property example for .NET | add PDF outline for each sheet in Excel to PDF conversion using Aspose.Cells | use reflection to set PdfSaveOptions.Bookmarks for older Aspose.Cells versions | C# code to export workbook to PDF with sheet bookmarks using Aspose.Cells
// Tags: Aspose.Cells PDF outline from worksheets | reflection set PDF bookmarks option Aspose.Cells | Excel workbook to PDF with sheet bookmarks | Aspose.Cells PDF bookmark feature .NET | enable PDF bookmarks Aspose.Cells save options

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// The sample loads an Excel workbook, uses reflection to enable the PdfSaveOptions.Bookmarks property when it exists, and saves the workbook as a PDF so that each worksheet appears as a bookmark in the generated document.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If the current Aspose.Cells version supports the Bookmarks property,
            // enable it via reflection to keep the code compatible with older versions.
            var bookmarksProp = typeof(PdfSaveOptions).GetProperty("Bookmarks");
            if (bookmarksProp != null && bookmarksProp.CanWrite)
            {
                bookmarksProp.SetValue(pdfOptions, true);
            }

            // Save the workbook as a PDF file with the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
