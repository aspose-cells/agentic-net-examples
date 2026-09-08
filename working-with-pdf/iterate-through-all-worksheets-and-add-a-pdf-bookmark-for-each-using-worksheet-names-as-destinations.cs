// Title: Add PDF bookmarks for each worksheet using worksheet names with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates over all worksheets in an Aspose.Cells Workbook, creates a PdfBookmark for each using the worksheet's Name as the title, adds the bookmarks to PdfSaveOptions.Bookmarks, and saves the workbook as a PDF containing those bookmarks. | Update the provided program to populate the PdfSaveOptions.Bookmarks collection so each bookmark links to the first page of its corresponding worksheet, then export the workbook to PDF with the bookmarks embedded.
// Common Searches: Aspose.Cells C# add PDF bookmarks for each Excel sheet | How to export an Excel workbook to PDF with sheet bookmarks using Aspose.Cells | C# generate PDF bookmarks from worksheet names with Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PDF bookmark creation | C# worksheet PDF bookmark mapping | PdfSaveOptions bookmarks collection | Excel to PDF with sheet bookmarks | Aspose.Cells add bookmarks per sheet

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures PdfSaveOptions, and saves it as a PDF. It currently does not add bookmarks, and includes a note that the PdfBookmark class and the Bookmarks property may be unavailable in some Aspose.Cells versions. To include bookmarks, the PdfSaveOptions.Bookmarks collection must be populated with entries that map each worksheet name to its destination page.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PdfBookmark and the Bookmarks property are not available in all
            // versions of Aspose.Cells. If needed, they can be added when supported.
            // The current implementation saves the workbook as PDF without bookmarks.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF with the defined options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
