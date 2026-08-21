// Title: PowerShell script to batch convert Excel files to HTML with AspNet.Cells ConversionUtility
// Description: A ready‑to‑run PowerShell example that loads the Aspose.Cells .NET assembly, scans a source folder (including subfolders) for *.xlsx, *.xls and *.xlsm files, creates a target directory, and uses ConversionUtility.Convert to generate matching .html files. The script logs successes, handles missing files and conversion errors, and can be scheduled for unattended execution.
// Keywords: PowerShell | Aspose.Cells | ConversionUtility | Excel to HTML | batch conversion | recursive folder scan | .NET library | automation script | bulk Excel conversion
// Common Searches: PowerShell batch convert Excel to HTML Aspose.Cells | How to use ConversionUtility.Convert in PowerShell | Convert multiple XLSX files to HTML via script | Recursive Excel to HTML conversion PowerShell example | Aspose.Cells PowerShell automation
// Developer Intent: Write a PowerShell script that imports Aspose.Cells, enumerates all Excel workbooks in a given directory tree, and calls ConversionUtility.Convert to produce HTML files with the same base names.
// Use Cases: Nightly job that publishes financial Excel reports as web‑ready HTML pages. | CI/CD step that transforms a library of spreadsheet templates into static HTML documentation. | Server‑side preview service that converts user‑uploaded Excel files to HTML without requiring Microsoft Office.
// AI Prompts: Generate a PowerShell script that loads the Aspose.Cells assembly and uses ConversionUtility.Convert to batch convert every .xlsx, .xls, and .xlsm file from C:\ExcelFiles to C:\HtmlOutput, preserving file names. | Create a PowerShell function Convert-ExcelFolderToHtml that accepts source and destination paths, creates the destination folder if needed, recursively finds Excel files, converts each to HTML with Aspose.Cells, and writes success or error messages to the console. | Provide PowerShell code that includes try/catch blocks for each conversion, logs detailed error information, and can be run unattended for large file sets.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExample
{
    // A ready‑to‑run PowerShell example that loads the Aspose.Cells .NET assembly, scans a source folder (including subfolders) for *.xlsx, *.xls and *.xlsm files, creates a target directory, and uses ConversionUtility.Convert to generate matching .html files. The script logs successes, handles missing files and conversion errors, and can be scheduled for unattended execution.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the folder that contains the Excel files to be converted
            string sourceFolder = @"C:\ExcelFiles";

            // Define the folder where the HTML files will be saved (create it if it doesn't exist)
            string targetFolder = @"C:\HtmlOutput";

            try
            {
                // Ensure the target folder exists
                if (!Directory.Exists(targetFolder))
                {
                    Directory.CreateDirectory(targetFolder);
                }

                // Get all Excel files (XLSX, XLS, XLSM) in the source folder recursively
                var excelFiles = Directory.EnumerateFiles(sourceFolder, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

                foreach (var filePath in excelFiles)
                {
                    // Verify the source file exists before attempting conversion
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    // Build the destination HTML file path (same name, .html extension)
                    string destFile = Path.Combine(targetFolder, Path.GetFileNameWithoutExtension(filePath) + ".html");

                    try
                    {
                        // Convert the Excel file to HTML using Aspose.Cells
                        ConversionUtility.Convert(filePath, destFile);
                        Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {destFile}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to convert {Path.GetFileName(filePath)}: {ex.Message}");
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
