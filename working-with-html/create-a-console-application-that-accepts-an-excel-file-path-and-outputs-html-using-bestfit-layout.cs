// Title: C# Console App: Convert Excel to HTML with Aspose.Cells using Best‑Fit (Normal) Layout
// Description: A lightweight .NET console utility that accepts an Excel file path, loads the workbook with Aspose.Cells, applies HtmlSaveOptions with HtmlLayoutMode.Normal for a best‑fit HTML layout, and saves the result as a .html file in the same folder. Includes argument validation, file‑existence checks, and exception handling.
// Keywords: Aspose.Cells | C# Excel to HTML conversion | HtmlLayoutMode.Normal | best‑fit HTML layout | HtmlSaveOptions example | .NET console Excel export | command line Excel to HTML
// Common Searches: convert excel workbook to html using aspose.cells c# | htmllayoutmode.normal example console app | c# command line tool to export excel as html | aspose.cells save workbook as html best fit layout | batch convert excel files to html with asp.net
// Developer Intent: Create a command‑line program that transforms an Excel workbook into a web‑ready HTML file using Aspose.Cells with the Normal (best‑fit) layout mode.
// Use Cases: Automate bulk conversion of Excel reports to HTML for web publishing. | Integrate Excel‑to‑HTML rendering into CI/CD pipelines to generate documentation previews. | Provide end‑users a quick, no‑Excel preview of worksheets via a simple executable.
// AI Prompts: Generate a C# console application that reads an Excel file path argument and saves it as HTML using Aspose.Cells with HtmlLayoutMode.Normal. | Extend the program to accept an optional output directory argument and write the HTML file there. | Add structured logging (e.g., to a file) for conversion successes and failures while preserving console output.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExcelToHtmlConverter
{
    // A lightweight .NET console utility that accepts an Excel file path, loads the workbook with Aspose.Cells, applies HtmlSaveOptions with HtmlLayoutMode.Normal for a best‑fit HTML layout, and saves the result as a .html file in the same folder. Includes argument validation, file‑existence checks, and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that an input file path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelToHtmlConverter <excel-file-path>");
                return;
            }

            string excelPath = args[0];

            // Check if the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: File not found - {excelPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified Excel file
                Workbook workbook = new Workbook(excelPath);

                // Create HTML save options (uses the default constructor rule)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Set layout mode to Normal (best‑fit layout similar to Excel)
                htmlOptions.LayoutMode = HtmlLayoutMode.Normal;

                // Determine output HTML file path (same folder, same name with .html extension)
                string outputPath = Path.ChangeExtension(excelPath, ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Conversion successful. HTML saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
