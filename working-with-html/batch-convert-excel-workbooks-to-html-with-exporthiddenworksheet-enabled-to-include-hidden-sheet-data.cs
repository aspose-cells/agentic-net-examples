// Title: Batch convert multiple Excel (.xls/.xlsx) workbooks to HTML with hidden worksheets using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert all Excel files in a folder to HTML including hidden sheets | How to export hidden worksheets to HTML with Aspose.Cells batch processing | Programmatically save Excel workbooks as HTML using ExportHiddenWorksheet option in .NET | Iterate through a directory of .xls and .xlsx files and generate HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet example for bulk conversion
// Tags: batch excel to html Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet | convert hidden worksheets to html c# | process multiple workbooks in directory | save workbook as html programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // The C# program scans a specified input folder for .xls and .xlsx files, loads each workbook with Aspose.Cells, sets HtmlSaveOptions.ExportHiddenWorksheet to true, and saves each workbook as an HTML file with the same base name in an output folder, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files to convert
            string inputFolder = @"C:\InputExcelFiles";
            // Folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtmlFiles";

            try
            {
                // Verify input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder not found: {inputFolder}");
                    return;
                }

                // Ensure the output directory exists
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                // Get all Excel files in the input folder (supports .xls and .xlsx)
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string filePath in excelFiles)
                {
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".xls" && extension != ".xlsx")
                        continue; // Skip non‑Excel files

                    // Verify the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Configure HTML save options to include hidden worksheets
                        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                        {
                            ExportHiddenWorksheet = true // Include hidden sheets in the output
                        };

                        // Determine output HTML file name (same base name as the Excel file)
                        string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the workbook as HTML with the specified options
                        workbook.Save(outputPath, saveOptions);

                        Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
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
