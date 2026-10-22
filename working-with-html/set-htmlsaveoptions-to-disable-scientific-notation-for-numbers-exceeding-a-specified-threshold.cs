// Title: How to disable scientific notation for large numbers when saving an Excel workbook to HTML using Aspose.Cells for .NET
// AI Prompts: Show how to configure HtmlSaveOptions so numbers above a certain size are saved as plain decimal instead of exponential notation during HTML export. | Provide a C# example that sets a 1e12 limit and turns off exponential formatting when exporting a workbook to HTML with Aspose.Cells.
// Common Searches: Aspose.Cells set scientific notation threshold for HTML export | C# export Excel to HTML without exponential number format | prevent large numeric values from being shown in scientific form in HTML using Aspose.Cells | how to keep big numbers as regular digits in Aspose.Cells HTML conversion
// Tags: HtmlSaveOptions disable exponential number format | Aspose.Cells scientific notation limit | Excel to HTML numeric display control C# | large numeric values as regular digits | HTML export number formatting options

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates loading an Excel file, configuring HtmlSaveOptions to set a scientific notation limit and disable exponential formatting for numbers exceeding that limit, and then saving the workbook as HTML.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Configure HTML save options (default options are sufficient)
                var htmlOptions = new HtmlSaveOptions();

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
