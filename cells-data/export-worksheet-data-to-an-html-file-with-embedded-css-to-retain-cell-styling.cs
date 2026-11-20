// Title: Export a single worksheet to an HTML file with inline CSS and Base64‑encoded images using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export active worksheet to HTML with embedded CSS | How to create a single HTML file from an Excel workbook using Base64 images in Aspose.Cells | Preserve Excel cell formatting when converting to HTML with Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 example in C# | Save only the active sheet as HTML using Aspose.Cells for .NET
// Tags: HtmlSaveOptions ExportImagesAsBase64 | Aspose.Cells export worksheet to HTML | inline CSS in HTML export Aspose.Cells | single HTML file from Excel C# | active worksheet only HTML conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The sample loads an existing .xlsx workbook, configures HtmlSaveOptions to embed images as Base64 strings and to export only the active worksheet, then saves the result as an HTML file that contains inline CSS, ensuring the original cell styling is retained.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourceFile = @"C:\Data\SampleWorkbook.xlsx";

            // Path for the output HTML file
            string outputFile = @"C:\Data\SampleWorkbook.html";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            try
            {
                // Load the workbook from the file system
                Workbook workbook = new Workbook(sourceFile);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Embed images as Base64 strings to keep everything in a single file
                    ExportImagesAsBase64 = true,

                    // Export only the active worksheet; set to false to export all worksheets
                    ExportActiveWorksheetOnly = true
                };

                // Save the workbook as an HTML file with the specified options
                workbook.Save(outputFile, saveOptions);

                Console.WriteLine("Worksheet exported to HTML successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }
    }
}
