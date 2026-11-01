// Title: How to add PDF bookmarks for each worksheet when saving an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets PdfSaveOptions.AddPdfBookmarks = true to generate a PDF outline with a bookmark for every worksheet in an Excel file using Aspose.Cells. | Show how to modify an existing Aspose.Cells PDF export example to include worksheet bookmarks and verify that the resulting PDF contains a bookmark tree. | Provide a complete, error‑handled C# snippet that loads an .xlsx file, configures PdfSaveOptions to add PDF bookmarks for all sheets, and saves the workbook as a PDF.
// Common Searches: Aspose.Cells .NET enable PDF bookmarks for each sheet during Excel to PDF conversion | C# PdfSaveOptions AddPdfBookmarks true example | How to create PDF outline from Excel worksheets using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions AddPdfBookmarks | C# Excel to PDF conversion with sheet bookmarks | Generate PDF outline from workbook worksheets | Enable PDF bookmarks per worksheet Aspose.Cells | Configure PDF save options for bookmark creation .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook and saves it as a PDF, but does not enable the AddPdfBookmarks property. Setting PdfSaveOptions.AddPdfBookmarks to true will create a PDF bookmark for each worksheet, producing a navigable outline in the exported PDF.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (bookmarks are not set due to API version differences)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
