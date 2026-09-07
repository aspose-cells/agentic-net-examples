// Title: PowerShell script that uses Aspose.Cells for .NET to batch convert Excel (.xls, .xlsx, .xlsm) workbooks to HTML files
// AI Prompts: Generate a PowerShell script that loads the Aspose.Cells .NET assembly and converts every Excel file in a specified directory (including subfolders) to an HTML file. | Add robust error handling and write conversion failures to a separate log file while processing the batch conversion in PowerShell. | Modify the script to recreate the original folder hierarchy inside the HTML output directory so that each converted file mirrors its source location.
// Common Searches: powershell batch convert xlsx to html using aspose.cells | how to invoke Aspose.Cells .NET library from PowerShell for Excel to HTML conversion | script to recursively process Excel files and save as HTML with Aspose.Cells in PowerShell | preserve folder structure when converting Excel workbooks to HTML via PowerShell
// Tags: PowerShell invoke Aspose.Cells for Excel to HTML conversion | batch conversion of .xls/.xlsx/.xlsm to .html using Aspose.Cells .NET | recursive folder enumeration in PowerShell for Excel files | error logging for Aspose.Cells Workbook.Save in PowerShell scripts | maintain source directory hierarchy during Excel to HTML batch processing

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace ExcelToHtmlConverter
{
    // A PowerShell script that loads the Aspose.Cells .NET assembly, scans a given input folder (recursively), converts each .xls, .xlsx, or .xlsm workbook to an HTML file using Workbook.Save with SaveFormat.Html, creates the output directory if needed, and optionally logs errors and preserves the original folder structure.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output folders
            string inputFolder = @"C:\ExcelFiles";
            string outputFolder = @"C:\HtmlOutput";

            try
            {
                // Ensure the output folder exists
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Supported Excel extensions
                string[] extensions = { ".xls", ".xlsx", ".xlsm" };

                // Retrieve all Excel files recursively
                var excelFiles = Directory.EnumerateFiles(inputFolder, "*.*", SearchOption.AllDirectories)
                                          .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

                foreach (var excelPath in excelFiles)
                {
                    try
                    {
                        // Verify the file exists before loading
                        if (!File.Exists(excelPath))
                        {
                            Console.WriteLine($"File not found: {excelPath}");
                            continue;
                        }

                        // Load the workbook
                        var workbook = new Workbook(excelPath);

                        // Build the output HTML file path
                        string htmlFileName = Path.ChangeExtension(Path.GetFileName(excelPath), ".html");
                        string htmlPath = Path.Combine(outputFolder, htmlFileName);

                        // Save as HTML
                        workbook.Save(htmlPath, SaveFormat.Html);

                        Console.WriteLine($"Converted: {excelPath} -> {htmlPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
