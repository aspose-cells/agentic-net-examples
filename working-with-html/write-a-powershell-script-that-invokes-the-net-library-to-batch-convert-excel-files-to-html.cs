// Title: PowerShell script for batch Excel‑to‑HTML conversion using Aspose.Cells .NET ConversionUtility
// Description: PowerShell example that loads the Aspose.Cells assembly and calls ConversionUtility.Convert to transform every .xlsx, .xls, and .xlsm file in a source folder (recursively) into HTML files, creates the output directory if needed, and logs successes and errors.
// Keywords: PowerShell | Aspose.Cells | ConversionUtility | Excel to HTML | batch conversion | recursive file enumeration | .NET library | automated report generation | script example | GitHub
// Common Searches: PowerShell batch convert Excel to HTML Aspose.Cells | How to call Aspose.Cells ConversionUtility from PowerShell | Recursive Excel to HTML conversion script .NET | PowerShell script for converting multiple .xlsx files to HTML | Error handling in PowerShell Excel to HTML conversion Aspose
// Developer Intent: Create a PowerShell script that leverages the Aspose.Cells .NET library to convert many Excel workbooks to HTML files in one operation.
// Use Cases: Generate HTML versions of daily Excel reports for web publishing. | Migrate a legacy Excel documentation library to static HTML pages. | Include batch Excel‑to‑HTML conversion in a CI/CD pipeline for documentation builds.
// AI Prompts: Write a PowerShell script that imports Aspose.Cells, creates the output folder, enumerates .xlsx/.xls/.xlsm files recursively, and uses ConversionUtility.Convert for each file with proper try/catch logging. | Provide PowerShell code that mirrors the given C# batch conversion logic, including directory creation, file validation, and per‑file exception handling. | Generate a PowerShell function named Convert-ExcelToHtml that accepts source and destination paths and performs bulk conversion using Aspose.Cells ConversionUtility, returning a summary of successes and failures.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells.Utility;

namespace ExcelToHtmlConverter
{
    // PowerShell example that loads the Aspose.Cells assembly and calls ConversionUtility.Convert to transform every .xlsx, .xls, and .xlsm file in a source folder (recursively) into HTML files, creates the output directory if needed, and logs successes and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the folder containing Excel files and the folder for HTML output
            string sourceFolder = @"C:\ExcelFiles";
            string outputFolder = @"C:\HtmlOutput";

            try
            {
                // Create the output folder if it does not exist
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Retrieve all Excel files (xlsx, xls, xlsm) recursively
                var excelFiles = Directory.EnumerateFiles(sourceFolder, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

                foreach (var excelPath in excelFiles)
                {
                    try
                    {
                        // Ensure the source file exists
                        if (!File.Exists(excelPath))
                        {
                            Console.WriteLine($"Source file not found: {excelPath}");
                            continue;
                        }

                        // Build the destination HTML file path (same name, .html extension)
                        string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                        string htmlPath = Path.Combine(outputFolder, htmlFileName);

                        // Perform the conversion using Aspose.Cells ConversionUtility
                        ConversionUtility.Convert(excelPath, htmlPath);

                        Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'");
                    }
                    catch (Exception ex)
                    {
                        // Handle conversion errors for individual files
                        Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle unexpected errors in the overall process
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
