// Title: Batch convert Excel workbooks to HTML with individual default fonts using Aspose.Cells for .NET
// Description: A C# console app that iterates over a list of Excel files, assigns a specific DefaultFontName via HtmlSaveOptions for each workbook, and saves them as HTML files. Includes validation, error handling, and logging of conversion results.
// Keywords: Aspose.Cells | C# HTML conversion | DefaultFontName | HtmlSaveOptions | batch Excel to HTML | multiple workbook export | set default font for HTML | Excel to HTML automation | Aspose.Cells example
// Common Searches: Aspose.Cells batch convert Excel to HTML C# | set default font per workbook HtmlSaveOptions | convert multiple Excel files to HTML with different fonts | C# export Excel as HTML Aspose.Cells default font | automate Excel to HTML conversion Aspose
// Developer Intent: Automatically export a collection of Excel workbooks to HTML, applying a unique default font to each file.
// Use Cases: Create web‑ready reports where each spreadsheet follows a distinct corporate typeface. | Generate documentation portals that require different font styles per source workbook. | Integrate into a CI/CD pipeline to batch‑process Excel assets into HTML with custom typography.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert a set of Excel files to HTML, assigning a specific DefaultFontName to each file. | Add comprehensive logging and exception handling to the batch HTML conversion script, writing results to a log file. | Refactor the example to output all HTML files into a dedicated folder while preserving original filenames.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchHtmlConversion
{
    // A C# console app that iterates over a list of Excel files, assigns a specific DefaultFontName via HtmlSaveOptions for each workbook, and saves them as HTML files. Includes validation, error handling, and logging of conversion results.
    class Program
    {
        static void Main()
        {
            // Define source workbook files and the corresponding default fonts
            string[] sourceFiles = { "Workbook1.xlsx", "Workbook2.xlsx", "Workbook3.xlsx" };
            string[] defaultFonts = { "Arial", "Courier New", "Times New Roman" };

            // Ensure the arrays have the same length
            if (sourceFiles.Length != defaultFonts.Length)
            {
                Console.WriteLine("The number of source files must match the number of fonts.");
                return;
            }

            // Process each workbook
            for (int i = 0; i < sourceFiles.Length; i++)
            {
                string sourcePath = sourceFiles[i];
                string fontName = defaultFonts[i];

                // Verify that the source file exists before attempting to load it
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: '{sourcePath}'. Skipping this file.");
                    continue;
                }

                try
                {
                    // Load the workbook (load rule)
                    Workbook workbook = new Workbook(sourcePath);

                    // Create HTML save options (create rule)
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        // Set a distinct default font for this conversion (feature rule)
                        DefaultFontName = fontName
                    };

                    // Determine the output HTML file name
                    string outputPath = Path.ChangeExtension(sourcePath, ".html");

                    // Save the workbook as HTML using the configured options (save rule)
                    workbook.Save(outputPath, saveOptions);

                    Console.WriteLine($"Converted '{sourcePath}' to HTML with default font '{fontName}'. Output: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{sourcePath}': {ex.Message}");
                }
            }
        }
    }
}
