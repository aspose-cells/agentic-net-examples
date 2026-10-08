// Title: PowerShell script that uses Aspose.Cells for .NET to batch convert .xls and .xlsx workbooks to HTML files
// AI Prompts: Write PowerShell code that loads the Aspose.Cells .NET assembly, iterates through all .xls and .xlsx files in a given directory (including subfolders), and saves each workbook as an .html file using Workbook.Save with SaveFormat.Html. | Create a PowerShell script that creates the target HTML output folder, includes try/catch blocks for folder creation and file conversion errors, and logs the conversion result for each Excel file.
// Common Searches: How to convert multiple Excel workbooks to HTML with PowerShell and Aspose.Cells | PowerShell batch processing of .xls and .xlsx files for HTML export using Aspose.Cells .NET | Script to recursively convert Excel files to HTML via Aspose.Cells from the command line | Automating Excel to HTML conversion in PowerShell with error handling | Using Aspose.Cells SaveFormat.Html in a PowerShell automation script
// Tags: Aspose.Cells PowerShell integration for HTML export | Excel workbook HTML rendering via .NET | PowerShell recursive directory traversal for file conversion | Error handling pattern for PowerShell and Aspose.Cells | Automated HTML generation from spreadsheets using Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlConverter
{
    // The script loads the Aspose.Cells .NET assembly, scans a specified input folder (including subfolders) for .xls and .xlsx files, creates an output directory, converts each workbook to an HTML file with SaveFormat.Html, and logs successes or errors for every file processed.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output folders
            string inputFolder = @"C:\ExcelFiles";
            string outputFolder = @"C:\HtmlOutput";

            try
            {
                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output folder '{outputFolder}': {ex.Message}");
                return;
            }

            // Get all .xls and .xlsx files recursively
            string[] excelFiles;
            try
            {
                excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.AllDirectories);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing input folder '{inputFolder}': {ex.Message}");
                return;
            }

            foreach (string excelPath in excelFiles)
            {
                string extension = Path.GetExtension(excelPath);
                if (!string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Skip non‑Excel files
                }

                // Build HTML output path
                string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Save as HTML
                    workbook.Save(htmlPath, SaveFormat.Html);

                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
    }
}
