// Title: C# – Export Excel to HTML with Percentage Column Widths (WidthScalable) using Aspose.Cells
// Description: This example loads an existing .xlsx workbook, sets HtmlSaveOptions.WidthScalable to true so column widths are saved as percentages, and saves the file as responsive HTML. It also includes basic file‑existence checking and exception handling.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | WidthScalable | percentage column widths | Excel to HTML | responsive HTML export | save workbook as HTML | scalable columns
// Common Searches: Aspose.Cells WidthScalable example | C# export Excel to HTML with percentage widths | How to make HTML columns responsive using Aspose.Cells | HtmlSaveOptions WidthScalable true C# | Convert XLSX to responsive HTML .NET
// Developer Intent: Generate an HTML document from an Excel workbook where column widths are expressed as percentages for responsive layouts.
// Use Cases: Create web‑ready reports that keep column proportions across desktop and mobile browsers. | Integrate Excel‑to‑HTML conversion into a REST API that returns mobile‑friendly tables. | Automate export of financial dashboards to HTML while preserving relative column sizing.
// AI Prompts: Write a C# snippet that loads an .xlsx file, sets HtmlSaveOptions.WidthScalable = true, and injects a custom CSS class into the exported HTML table. | Explain how WidthScalable changes the generated HTML and recommend additional HtmlSaveOptions for full responsiveness, such as EmbedImages and ExportActiveWorksheetOnly. | Provide best‑practice error handling for checking the input file, logging exceptions, and returning user‑friendly messages when exporting Excel to HTML with scalable column widths.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example loads an existing .xlsx workbook, sets HtmlSaveOptions.WidthScalable to true so column widths are saved as percentages, and saves the file as responsive HTML. It also includes basic file‑existence checking and exception handling.
    public class WidthScalableHtmlExport
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options to use scalable column widths (percentage)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    WidthScalable = true
                };

                // Save the workbook as an HTML file
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WidthScalableHtmlExport.Run();
        }
    }
}
