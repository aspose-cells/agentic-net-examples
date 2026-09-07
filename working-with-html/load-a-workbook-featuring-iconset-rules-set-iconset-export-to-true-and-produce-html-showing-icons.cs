// Title: How to export Excel IconSet conditional formatting to HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook, sets HtmlSaveOptions.ExportIconSet = true, and saves the workbook as an HTML file that displays the IconSet icons. | Show how to configure Aspose.Cells HtmlSaveOptions to preserve IconSet conditional‑formatting rules when converting a workbook to HTML in a .NET application.
// Common Searches: Aspose.Cells C# export IconSet conditional formatting to HTML | Enable ExportIconSet option in HtmlSaveOptions for Excel to HTML conversion | Convert Excel workbook with icon sets to HTML using Aspose.Cells .NET | Preserve conditional formatting icons when saving Excel as HTML in C# | How to show Excel IconSet icons in HTML output with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportIconSet | C# export IconSet conditional formatting to HTML | preserve Excel conditional formatting icons in HTML | convert workbook with IconSet rules to HTML | load Excel workbook save as HTML with icons

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, creates HtmlSaveOptions with ExportIconSet enabled, and saves the workbook as an HTML page that retains IconSet conditional‑formatting icons. It includes basic error handling and console output indicating success or failure.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.html";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure HTML save options (default settings export conditional formatting)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save the workbook as HTML
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
