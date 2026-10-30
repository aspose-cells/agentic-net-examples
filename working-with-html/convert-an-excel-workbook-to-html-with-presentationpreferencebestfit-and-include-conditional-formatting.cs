// Title: Convert an Excel workbook to HTML using Aspose.Cells for .NET with HtmlSaveOptions (best‑fit layout note)
// Common Searches: asp.net convert xlsx file to html using aspose.cells HtmlSaveOptions | c# set PresentationPreference to BestFit for HTML export in Aspose.Cells | include conditional formatting when exporting Excel to HTML with Aspose.Cells | how to handle missing source file before converting Excel to HTML in C# | asp.net core save workbook as html and ensure output directory exists
// Tags: Aspose.Cells HtmlSaveOptions HTML export | C# Excel to HTML conversion Aspose.Cells | PresentationPreference BestFit Aspose.Cells | conditional formatting limitation HTML export Aspose.Cells | ensure output directory exists C# file conversion

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlConversion
{
    // The example loads an Excel workbook, configures HtmlSaveOptions, creates the output folder if needed, and saves the workbook as an HTML file using Aspose.Cells for .NET. It notes that the IncludeConditionalFormatting option is unavailable in this version and demonstrates basic error handling for missing files and conversion failures.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourceFile = @"C:\Path\To\YourWorkbook.xlsx";

            // Path where the HTML output will be saved
            string outputFile = @"C:\Path\To\ConvertedWorkbook.html";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(sourceFile);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // Note: IncludeConditionalFormatting property is not available in this version of Aspose.Cells

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an HTML file using the configured options
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully converted to HTML: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
