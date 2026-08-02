// Title: Batch convert Excel workbooks to HTML with hidden sheets using Aspose.Cells (C#)
// Description: A C# console app that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, enables HtmlSaveOptions.ExportHiddenWorksheet, and saves the result as HTML in a target directory while handling missing files and errors.
// Keywords: Aspose.Cells | C# | batch Excel to HTML | ExportHiddenWorksheet | hidden worksheets | HTML conversion .NET | folder processing | automation example | Excel to web report | Aspose.Cells HTMLSaveOptions
// Common Searches: Aspose.Cells batch convert Excel to HTML | C# export hidden worksheets to HTML | HtmlSaveOptions ExportHiddenWorksheet example | convert all xlsx files in a folder to html | Aspose.Cells hide sheet HTML output
// Developer Intent: Convert every .xlsx file in a specified directory to an HTML page that includes data from hidden worksheets.
// Use Cases: Publish a set of Excel templates with hidden calculation sheets as web‑ready reports. | Automate nightly processing of uploaded workbooks, preserving hidden data for intranet viewers. | Create a CI/CD step that transforms financial models into static HTML for documentation.
// AI Prompts: Generate C# code that batch converts Excel files in a directory to HTML while including hidden worksheets using Aspose.Cells. | Show how to modify the script to output each worksheet as a separate HTML file and keep hidden sheet content. | Explain how to add structured logging and retry logic for large‑scale batch conversions with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // A C# console app that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, enables HtmlSaveOptions.ExportHiddenWorksheet, and saves the result as HTML in a target directory while handling missing files and errors.
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
                Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder
            foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Configure HTML save options to include hidden worksheets
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExportHiddenWorksheet = true // Include hidden sheet data
                    };

                    // Determine output HTML file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, htmlOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
