// Title: Batch Convert Excel Files to Inline‑Style HTML (Disable CSS) with Aspose.Cells for .NET
// Description: A C# utility that scans a folder for .xls and .xlsx workbooks, loads each with Aspose.Cells, and saves them as lightweight HTML files using HtmlSaveOptions.DisableCss to embed styles inline. Includes folder validation and per‑file error handling.
// Keywords: Aspose.Cells batch HTML export | disable CSS Aspose.Cells | Excel to HTML .NET | HtmlSaveOptions DisableCss example | inline style HTML from Excel | convert multiple workbooks to HTML | C# Aspose.Cells folder processing
// Common Searches: convert all Excel files in a directory to HTML using Aspose.Cells | Aspose.Cells HtmlSaveOptions.DisableCss batch conversion | C# code to export Excel workbooks to HTML without external CSS | how to generate lightweight HTML from Excel with Aspose
// Developer Intent: Transform every Excel workbook in a specified directory into HTML files that use only inline styles.
// Use Cases: Publish a collection of Excel‑based reports on a website without extra CSS requests. | Automate documentation generation in CI/CD pipelines, converting spreadsheets to self‑contained HTML pages. | Provide fast preview of user‑uploaded spreadsheets on a server, minimizing load time by avoiding separate style sheets.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert .xls/.xlsx files to HTML with DisableCss enabled, including folder checks and exception handling. | Explain how to modify the program to also export embedded chart images while keeping CSS disabled for the HTML output. | Show how to log each conversion (input path, output path, success/failure) to a CSV file instead of writing to the console.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// A C# utility that scans a folder for .xls and .xlsx workbooks, loads each with Aspose.Cells, and saves them as lightweight HTML files using HtmlSaveOptions.DisableCss to embed styles inline. Includes folder validation and per‑file error handling.
class Program
{
    static void Main(string[] args)
    {
        // Input folder containing Excel files (default: "Input")
        string inputFolder = args.Length > 0 ? args[0] : "Input";

        // Output folder for generated HTML files (default: "Output")
        string outputFolder = args.Length > 1 ? args[1] : "Output";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder \"{inputFolder}\" does not exist.");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        try
        {
            // Get all Excel files in the input folder (xls and xlsx)
            var excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                                      .Where(f => f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                                  f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Load the workbook from the Excel file
                    Workbook workbook = new Workbook(excelPath);

                    // Configure HTML save options to disable external CSS (use inline styles only)
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        DisableCss = true // lightweight HTML without separate CSS files
                    };

                    // Determine the output HTML file path
                    string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                    string htmlPath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Converted \"{Path.GetFileName(excelPath)}\" to \"{htmlFileName}\".");
                }
                catch (Exception exFile)
                {
                    Console.WriteLine($"Error processing file \"{excelPath}\": {exFile.Message}");
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
