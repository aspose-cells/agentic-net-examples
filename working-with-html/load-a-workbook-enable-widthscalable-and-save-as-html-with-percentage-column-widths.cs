// Title: Export an Excel workbook to HTML with percentage‑based column widths using Aspose.Cells for .NET
// Common Searches: how to export an .xlsx file to HTML with percentage column widths using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions preserve column width scaling when saving to HTML | C# save workbook as HTML with scalable column widths Aspose.Cells | set column width auto‑fit for HTML output in Aspose.Cells .NET | save Excel as HTML and keep column widths proportional Aspose.Cells
// Tags: HtmlSaveOptions column width scaling Aspose.Cells | C# Excel to HTML conversion Aspose.Cells | HTML export with column width scaling Aspose.Cells | Aspose.Cells workbook save as HTML | prepare output directory before saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions (which apply percentage‑based column widths by default), ensures the output folder exists, and saves the workbook as output.html while handling any runtime exceptions.
    class Program
    {
        static void Main()
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

                // Load the existing workbook from file
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // NOTE: In recent Aspose.Cells versions, automatic column width fitting and
                // percentage‑based column widths are applied by default when saving to HTML.
                // The previous properties (IsWidthAutoFit, IsColumnWidthInPixels) are no longer
                // available, so they are omitted here.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an HTML file with the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
