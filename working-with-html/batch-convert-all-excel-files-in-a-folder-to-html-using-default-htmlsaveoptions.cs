// Title: C# – Batch Convert Excel Files in a Folder to HTML with Aspose.Cells Default HtmlSaveOptions
// Description: A C# console example that scans a directory, loads each .xls, .xlsx, .xlsm or .csv workbook, and saves it as an .html file using Aspose.Cells ConversionUtility with default HtmlSaveOptions, including error handling and logging.
// Keywords: Aspose.Cells | C# batch Excel to HTML | convert folder of Excel files to HTML | HtmlSaveOptions default | ConversionUtility | Excel to HTML programmatically | Aspose.Cells example GitHub | Excel workbook HTML export
// Common Searches: batch convert excel to html c# asp.net | aspocells convert all xlsx files in a directory to html | c# example using ConversionUtility to save workbook as html | how to export multiple Excel workbooks to html with Aspose.Cells | default htmlsaveoptions aspocells sample code
// Developer Intent: Programmatically convert every Excel workbook in a specified folder to an HTML file using Aspose.Cells with default save settings.
// Use Cases: Generate web‑ready reports from a collection of spreadsheets on a server. | Provide HTML previews for uploaded Excel files in a web portal. | Migrate legacy Excel documentation to static HTML pages for easy distribution. | Automate batch export of financial models to HTML for publishing.
// AI Prompts: Write a C# console program that iterates through a directory and uses Aspose.Cells ConversionUtility to convert each .xls, .xlsx, .xlsm, or .csv file to HTML with default HtmlSaveOptions, including logging and exception handling. | Show an alternative approach using Workbook.Save for batch conversion of Excel files to HTML in C#. | Explain how to modify HtmlSaveOptions to embed images as base64 while keeping the batch conversion loop. | Provide a PowerShell script that calls the compiled C# batch converter for scheduled execution.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchExcelToHtml
{
    // A C# console example that scans a directory, loads each .xls, .xlsx, .xlsm or .csv workbook, and saves it as an .html file using Aspose.Cells ConversionUtility with default HtmlSaveOptions, including error handling and logging.
    class Program
    {
        static void Main()
        {
            // Folder containing the Excel files to be converted
            string sourceFolder = @"C:\ExcelFiles";

            // Ensure the folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Get all Excel files in the folder (common extensions)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".csv")
                    continue; // Skip non‑Excel files

                try
                {
                    // Destination HTML file path (same name, .html extension)
                    string destPath = Path.ChangeExtension(filePath, ".html");

                    // Load options – default constructor (no special settings)
                    LoadOptions loadOptions = new LoadOptions();

                    // Save options – default HtmlSaveOptions (as required)
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                    // Perform the conversion using the provided ConversionUtility rule
                    ConversionUtility.Convert(filePath, loadOptions, destPath, saveOptions);

                    Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {Path.GetFileName(destPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
