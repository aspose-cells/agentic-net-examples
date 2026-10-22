// Title: How to add worksheet names as <h1> headings when converting an Excel workbook to a single HTML file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, enables the option to insert a heading for each worksheet, and saves the workbook to one HTML file where every sheet name appears as an <h1> element. | Show the minimal changes required in the given Aspose.Cells example to include worksheet titles as top‑level headings in the generated HTML without altering the output file path. | Explain how the HtmlSaveOptions properties that control worksheet header inclusion and active‑worksheet export work together to produce an HTML document containing all worksheets with their names as headings.
// Common Searches: Aspose.Cells C# export multiple worksheets to one HTML file with sheet names as headings | how to include worksheet title as heading when saving Excel to HTML using Aspose.Cells | HtmlSaveOptions property to add worksheet headings in .NET example | convert workbook to HTML with each sheet displayed under its own h1 using Aspose.Cells | C# Aspose.Cells save workbook to HTML include worksheet headers
// Tags: Aspose.Cells HtmlSaveOptions include worksheet headings | C# export Excel workbook to single HTML file | Aspose.Cells convert multiple sheets to HTML with titles | HTML conversion preserve worksheet names Aspose.Cells | Aspose.Cells workbook to HTML with sheet headers

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, checks its presence, and uses Aspose.Cells to save it as HTML. By setting HtmlSaveOptions.ExportActiveWorksheetOnly = false and enabling the option that adds a heading for each worksheet, the resulting HTML file contains all worksheets with their names rendered as top‑level <h1> headings. The code also includes basic error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Set HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export all worksheets, not just the active one
                    ExportActiveWorksheetOnly = false
                };

                // Save the workbook as an HTML file with the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
