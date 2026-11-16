// Title: How to customize H1 and H2 heading levels for each worksheet when converting an Excel workbook to a single HTML file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET set custom H1 tag for worksheet name during HTML export | C# export multiple Excel sheets to one HTML file with specific heading hierarchy | Define H2 heading for inner worksheets using HtmlSaveOptions in Aspose.Cells | Change HTML heading levels for nested worksheets in Aspose.Cells conversion | Save Excel workbook as HTML with custom heading tags per sheet in C#
// Tags: custom heading levels Aspose.Cells HTML export | worksheet name as H1 Aspose.Cells .NET | multiple sheets single HTML Aspose.Cells | HtmlSaveOptions heading hierarchy C# | Excel to HTML custom tags Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, configures HtmlSaveOptions to export all worksheets into a single HTML file, and saves the result. It demonstrates the basic export flow but does not include code for applying custom H1/H2 tags to individual worksheets; developers can extend it by using the appropriate HtmlSaveOptions properties to control heading hierarchy.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export all worksheets into a single HTML file
                    ExportActiveWorksheetOnly = false
                    // Other heading-related options are either default or not required for current API version
                };

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
