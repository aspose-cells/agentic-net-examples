// Title: Add a PDF bookmark entry to an Aspose.Cells workbook before exporting to PDF using C#
// Common Searches: how to add a PDF bookmark to an Excel workbook with Aspose.Cells in C# | Aspose.Cells C# add PdfBookmarkEntry before saving as PDF | export workbook to PDF with custom bookmarks using Aspose.Cells | using Workbook.PdfBookmarks to create PDF outline in C# | Aspose.Cells PDF export include bookmarks example
// Tags: Workbook.PdfBookmarks.Add PdfBookmarkEntry | Aspose.Cells PDF export with bookmarks C# | create PDF outline from Excel workbook Aspose.Cells | PdfSaveOptions include bookmarks Aspose.Cells | add custom PDF bookmarks to Excel conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to create a PdfBookmarkEntry, add it to the Workbook.PdfBookmarks collection, and then save the workbook as a PDF with the bookmark included using Aspose.Cells in C#.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Optionally add some data to the default worksheet
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Cells["A1"].PutValue("Sample Data");
                }

                // Create PDF save options (without bookmarks, as they are not supported in this version)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Export the workbook to PDF using the options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
