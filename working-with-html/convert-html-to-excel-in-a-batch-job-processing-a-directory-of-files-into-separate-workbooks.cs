// Title: Batch convert a folder of HTML files to individual Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that enumerates all *.html files in a specified directory, loads each with Aspose.Cells HtmlLoadOptions, and saves them as separate .xlsx workbooks. | Write code that creates the output folder, validates the input path, and logs both successful conversions and any failures for each HTML file. | Implement per‑file exception handling inside a conversion loop that transforms HTML to Excel using Aspose.Cells and outputs the source‑to‑target file mapping.
// Common Searches: aspnet batch convert html files to xlsx using Aspose.Cells | c# console app process a directory of html and generate separate excel workbooks | how to use HtmlLoadOptions for multiple html to excel conversion in .NET | error handling for bulk html to excel conversion with Aspose.Cells
// Tags: Aspose.Cells HTML to XLSX batch conversion | C# directory enumeration for file format conversion | HtmlLoadOptions usage in bulk Excel export | per‑file exception handling in Aspose.Cells conversion | save workbook as SaveFormat.Xlsx in loop

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelBatchApp
{
    // // This C# console application scans a given input folder for *.html files, loads each file into an Aspose.Cells Workbook using HtmlLoadOptions, and saves the result as an individual .xlsx workbook in an output folder. It ensures the input folder exists, creates the output folder if needed, and provides per‑file error handling with console logging of conversion outcomes.
    class HtmlToExcelBatch
    {
        static void Main(string[] args)
        {
            // Directory containing the source HTML files
            string inputDirectory = @"C:\InputHtml";

            // Directory where the generated Excel workbooks will be saved
            string outputDirectory = @"C:\OutputExcel";

            try
            {
                // Verify input directory exists
                if (!Directory.Exists(inputDirectory))
                {
                    Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Retrieve all HTML files from the input directory (non‑recursive)
                string[] htmlFiles = Directory.GetFiles(inputDirectory, "*.html", SearchOption.TopDirectoryOnly);

                foreach (string htmlFilePath in htmlFiles)
                {
                    try
                    {
                        // Load the HTML file into an Aspose.Cells Workbook
                        Workbook workbook = new Workbook(htmlFilePath, new HtmlLoadOptions());

                        // Build the output Excel file path (same base name, .xlsx extension)
                        string baseName = Path.GetFileNameWithoutExtension(htmlFilePath);
                        string excelFilePath = Path.Combine(outputDirectory, baseName + ".xlsx");

                        // Save the workbook as an Excel file
                        workbook.Save(excelFilePath, SaveFormat.Xlsx);

                        Console.WriteLine($"Converted: {htmlFilePath} -> {excelFilePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to convert '{htmlFilePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
