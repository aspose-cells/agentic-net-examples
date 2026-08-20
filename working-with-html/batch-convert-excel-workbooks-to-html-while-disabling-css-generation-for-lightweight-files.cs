// Title: Batch convert Excel workbooks to HTML with inline styles (DisableCss) – Aspose.Cells for .NET
// Description: C# utility that scans a folder for Excel and CSV files, loads each workbook with Aspose.Cells, and saves it as HTML using HtmlSaveOptions.DisableCss = true, producing lightweight pages with only inline styling.
// Keywords: Aspose.Cells | C# Excel to HTML | batch conversion | DisableCss | inline styles | no external CSS | HTML preview | GitHub example | code snippet | convert folder Excel files
// Common Searches: Aspose.Cells batch convert Excel to HTML without CSS | C# convert multiple .xlsx files to HTML inline styles | HtmlSaveOptions DisableCss example for .NET | How to export a folder of Excel workbooks to HTML using Aspose | Convert CSV and Excel to lightweight HTML with Aspose.Cells
// Developer Intent: Automatically transform every supported Excel or CSV file in a directory into an HTML document that uses only inline CSS, eliminating external stylesheet files.
// Use Cases: Generate compact HTML previews of uploaded spreadsheets for web portals without managing separate CSS files. | Create searchable, SEO‑friendly HTML archives of legacy Excel reports while keeping file size minimal. | Automate conversion of data sets (Excel/CSV) into HTML snippets for newsletters, documentation, or email content.
// AI Prompts: Write a reusable C# method that takes input and output folder paths and uses Aspose.Cells to batch‑convert all Excel files to HTML with DisableCss enabled, including logging and error handling. | Show how to extend the batch conversion to also export chart images as separate files while still disabling CSS generation. | Explain how to modify HtmlSaveOptions to embed custom fonts and control page layout when converting many workbooks to HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // C# utility that scans a folder for Excel and CSV files, loads each workbook with Aspose.Cells, and saves it as HTML using HtmlSaveOptions.DisableCss = true, producing lightweight pages with only inline styling.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel files
            string inputFolder = @"C:\InputExcelFiles";

            // Output folder for generated HTML files
            string outputFolder = @"C:\OutputHtmlFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

            try
            {
                // Iterate through each file in the input folder
                foreach (string filePath in Directory.GetFiles(inputFolder))
                {
                    // Process only supported Excel files
                    if (Array.Exists(extensions, ext => ext.Equals(Path.GetExtension(filePath), StringComparison.OrdinalIgnoreCase)))
                    {
                        try
                        {
                            // Ensure the file still exists before loading
                            if (!File.Exists(filePath))
                            {
                                Console.WriteLine($"File not found (skipped): {filePath}");
                                continue;
                            }

                            // Load the workbook from the source file
                            Workbook workbook = new Workbook(filePath);

                            // Create HTML save options and disable CSS generation (use inline styles only)
                            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                            {
                                DisableCss = true // Inline styles, no external CSS
                            };

                            // Determine output HTML file name (same base name with .html extension)
                            string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                            string outputPath = Path.Combine(outputFolder, outputFileName);

                            // Save the workbook as HTML using the configured options
                            workbook.Save(outputPath, htmlOptions);

                            Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML successfully.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error converting '{Path.GetFileName(filePath)}': {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
