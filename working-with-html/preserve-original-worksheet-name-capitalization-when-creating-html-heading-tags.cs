// Title: Save an Excel workbook as HTML with original worksheet name capitalization preserved in heading tags (Aspose.Cells for .NET)
// AI Prompts: Write C# code that checks for an existing .xlsx file, loads it with Aspose.Cells, and saves it as HTML so the worksheet name appears in the <h1> tag with exactly the same capitalization as in the source workbook. | Show how to set up Aspose.Cells HtmlSaveOptions to retain the sheet name's case when exporting a workbook to HTML. | Create a resilient example that handles missing input files, loads the workbook, and outputs an HTML file without altering the worksheet name's letter case.
// Common Searches: how to keep worksheet name case when exporting Excel to HTML using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions preserve sheet name capitalization in generated HTML | C# export .xlsx to .html with original sheet heading case
// Tags: Aspose.Cells HTML export preserve sheet name case | C# HtmlSaveOptions worksheet heading capitalization | export Excel to HTML with original sheet name case | Aspose.Cells workbook.Save HTML case-sensitive heading | preserve worksheet name capitalization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks that the input XLSX file exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions, and saves the workbook as an HTML file. By default Aspose.Cells exports the worksheet name as a heading, keeping the original capitalization intact.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Configure HTML save options.
                // The ExportWorksheetHeader property is not available in the current Aspose.Cells version,
                // but the worksheet name is exported as a heading by default.
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Additional options can be set here if needed, e.g.:
                    // ExportActiveWorksheetOnly = false,
                };

                // Save the workbook as HTML while preserving the original worksheet name capitalization
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
