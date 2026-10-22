// Title: Convert an Excel workbook to a self‑contained HTML file with active hyperlinks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file with Aspose.Cells and saves it as an HTML document while keeping all hyperlinks clickable. | Show how to set HtmlSaveOptions to embed worksheet images as Base64 strings and export every worksheet in the workbook. | Add error handling that verifies the source file exists, creates the destination folder if needed, and logs any conversion exceptions.
// Common Searches: Aspose.Cells .NET export Excel to HTML preserving hyperlinks | C# generate self‑contained HTML from .xlsx with embedded images | How to keep Excel cell links active when converting to HTML using Aspose | HtmlSaveOptions ExportActiveWorksheetOnly false example | Save workbook as HTML with Base64 images Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | Aspose.Cells embed images Base64 HTML | Aspose.Cells preserve hyperlinks HTML export | C# convert Excel to self-contained HTML | Aspose.Cells create output directory before save

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlExport
{
    // The program loads an Excel workbook, configures HtmlSaveOptions to export all worksheets and embed images as Base64, ensures the output folder exists, and saves the workbook as a self‑contained HTML file where all original hyperlinks remain functional, with error handling for missing files and conversion failures.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = @"C:\Input\SampleWorkbook.xlsx";

            // Path for the generated HTML file
            string htmlPath = @"C:\Output\SampleWorkbook.html";

            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            try
            {
                // Load the workbook from the file system
                Workbook workbook = new Workbook(excelPath);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(htmlPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export all worksheets (set to true to export only the active sheet)
                    ExportActiveWorksheetOnly = false,

                    // Embed images as Base64 strings to keep the HTML self‑contained
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine("Excel file has been successfully exported to HTML with functional hyperlinks.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
