// Title: Convert multiple HTML files to individual Excel workbooks in a batch using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a given directory for *.html files, loads each file with Aspose.Cells HtmlLoadOptions, and saves it as a separate .xlsx workbook in an output folder. | Implement per‑file try/catch logic so that if converting one HTML file fails, the batch process continues with the remaining files. | Extend the program to copy the original HTML file's last‑modified timestamp to the generated Excel workbook after saving.
// Common Searches: aspnet batch convert html files to xlsx using Aspose.Cells | c# console app to convert a folder of html pages to separate Excel workbooks | how to process multiple html files with Aspose.Cells HtmlLoadOptions in .NET | bulk html to excel conversion with error handling in C# | preserve original file timestamps when exporting html to xlsx with Aspose.Cells
// Tags: Aspose.Cells batch HTML to XLSX conversion | C# HtmlLoadOptions workbook loading | directory iteration for file format conversion | per‑file exception handling in bulk conversion | preserve file timestamps during Excel export

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelBatch
{
    // The sample enumerates all *.html files in a specified input folder, loads each file into an Aspose.Cells Workbook using HtmlLoadOptions, and saves it as an .xlsx workbook with the same base name in an output folder. It creates the output directory if needed, logs per‑file errors, and continues processing despite individual failures.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the input directory containing HTML files
            string inputDirectory = @"C:\InputHtmlFiles";

            // Define the output directory where Excel workbooks will be saved
            string outputDirectory = @"C:\OutputExcelFiles";

            try
            {
                // Ensure the input directory exists; if not, inform the user and exit
                if (!Directory.Exists(inputDirectory))
                {
                    Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                    return;
                }

                // Ensure the output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // Get all HTML files in the input directory (top level only)
                string[] htmlFiles = Directory.GetFiles(inputDirectory, "*.html", SearchOption.TopDirectoryOnly);

                foreach (string htmlFilePath in htmlFiles)
                {
                    try
                    {
                        // Load the HTML file into an Aspose.Cells Workbook
                        Workbook workbook = new Workbook(htmlFilePath, new HtmlLoadOptions());

                        // Determine the output file name (same as input but with .xlsx extension)
                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(htmlFilePath);
                        string outputFilePath = Path.Combine(outputDirectory, fileNameWithoutExtension + ".xlsx");

                        // Save the workbook as an Excel file
                        workbook.Save(outputFilePath, SaveFormat.Xlsx);

                        Console.WriteLine($"Converted: {htmlFilePath} -> {outputFilePath}");
                    }
                    catch (Exception ex)
                    {
                        // Log any errors for the current file and continue processing others
                        Console.WriteLine($"Error processing file '{htmlFilePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
