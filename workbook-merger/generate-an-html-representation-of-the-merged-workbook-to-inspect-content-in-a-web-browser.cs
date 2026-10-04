// Title: Generate an HTML view of a merged Excel workbook with grid lines and Base64‑embedded images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a merged .xlsx workbook with Aspose.Cells, sets HtmlSaveOptions to export grid lines and embed images as Base64, and saves the result as an .html file. | Create a .NET snippet that verifies the input file exists, opens it with Aspose.Cells, configures ExportGridLines = true and ExportImagesAsBase64 = true, then writes the HTML output to a target path.
// Common Searches: asp.net convert merged workbook.xlsx to html with grid lines using Aspose.Cells | how to export Excel file as html with embedded images base64 in C# | Aspose.Cells HtmlSaveOptions ExportGridLines true example | save merged Excel workbook as web‑ready HTML C# Aspose | C# code to generate HTML preview of an Excel workbook with images embedded
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | Aspose.Cells ExportImagesAsBase64 | C# export Excel to HTML | merged workbook HTML preview Aspose.Cells | grid lines in HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The program checks for a mergedWorkbook.xlsx file, loads it with Aspose.Cells, configures HtmlSaveOptions to include grid lines and embed images as Base64 strings, and saves the workbook as mergedWorkbook.html while handling errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (merged workbook)
            string inputPath = "mergedWorkbook.xlsx";

            // Path where the HTML representation will be saved
            string outputPath = "mergedWorkbook.html";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                    return;
                }

                // Load the workbook from the file system
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export grid lines
                    ExportGridLines = true,

                    // Embed images as base64 strings (true) or save them as separate files (false)
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file using the specified options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML representation saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
