// Title: Generate a single HTML file containing all worksheets from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that loads an .xlsx workbook and saves it as one HTML document with every worksheet included using Aspose.Cells. | Show how to configure HtmlSaveOptions.ExportActiveWorksheetOnly to false so that all sheets are exported into a combined HTML file. | Adapt an existing Aspose.Cells example to verify the input file, set the proper HTML save options, and produce a single HTML output with every worksheet.
// Common Searches: Aspose.Cells C# export entire workbook to single HTML page | How to include all worksheets when saving Excel as HTML with Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly false example in .NET | Combine multiple Excel sheets into one HTML file using Aspose.Cells | Save Excel workbook as combined HTML using C# Aspose.Cells library
// Tags: Aspose.Cells HTML export all worksheets | HtmlSaveOptions ExportActiveWorksheetOnly false | C# save workbook as combined HTML | Aspose.Cells multiple sheets to HTML | export Excel to single HTML .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, sets HtmlSaveOptions.ExportActiveWorksheetOnly to false to include every worksheet, and saves the workbook as a single HTML file while handling errors and reporting the outcome.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export all worksheets (default behavior). Setting this explicitly for clarity.
                    ExportActiveWorksheetOnly = false
                };

                // Save the workbook as a single HTML file
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
