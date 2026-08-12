// Title: C# – Extract Excel Files from a ZIP Archive and Convert to HTML with Aspose.Cells
// Description: The script checks for a ZIP file, extracts its contents to a temporary directory, scans for supported Excel formats (.xls, .xlsx, .xlsm, .xlsb, .csv), converts each workbook to an HTML file using Aspose.Cells.Utility.ConversionUtility, saves the results to a target folder, and optionally deletes the extraction folder.
// Keywords: Aspose.Cells | C# Excel to HTML conversion | zip extraction C# | batch spreadsheet conversion | ConversionUtility.Convert | temporary folder cleanup | supported Excel extensions
// Common Searches: How to convert Excel files inside a zip to HTML with Aspose.Cells C# | Aspose.Cells batch conversion from zip archive | C# extract zip and convert multiple .xlsx to HTML | Convert .xls/.xlsx to HTML using ConversionUtility | Delete temporary extraction folder after processing Excel files
// Developer Intent: Extract every Excel workbook from a ZIP archive and generate a corresponding HTML file using Aspose.Cells in a C# application.
// Use Cases: Automate the transformation of user‑uploaded ZIP packages of spreadsheets into web‑ready HTML reports for a portal. | Run a scheduled Windows service that processes a shared folder of zipped workbooks and publishes HTML versions for intranet browsing. | Integrate into a document‑management workflow to turn archived Excel files into searchable HTML pages without manual intervention.
// AI Prompts: Write C# code that extracts Excel files from a zip archive, converts each to HTML with Aspose.Cells ConversionUtility, and includes comprehensive error handling. | Refactor the script to use asynchronous I/O and streams so that a temporary extraction folder is not required. | Explain how to customize the HTML output (styles, images, encoding) when using ConversionUtility.Convert. | Generate a PowerShell wrapper that calls the C# executable to process zip files in a CI/CD pipeline.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells.Utility;

namespace AsposeCellsZipToHtml
{
    // The script checks for a ZIP file, extracts its contents to a temporary directory, scans for supported Excel formats (.xls, .xlsx, .xlsm, .xlsb, .csv), converts each workbook to an HTML file using Aspose.Cells.Utility.ConversionUtility, saves the results to a target folder, and optionally deletes the extraction folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the zip archive containing Excel files
            string zipPath = @"C:\Input\excel_files.zip";

            // Verify that the zip file exists before proceeding
            if (!File.Exists(zipPath))
            {
                Console.WriteLine($"Error: Zip file not found at '{zipPath}'.");
                return;
            }

            // Temporary folder to extract the Excel files
            string extractFolder = Path.Combine(Path.GetTempPath(), "ExcelExtracted");
            Directory.CreateDirectory(extractFolder);

            // Folder where the resulting HTML files will be saved
            string htmlOutputFolder = @"C:\Output\HtmlFiles";
            Directory.CreateDirectory(htmlOutputFolder);

            // Extract all entries from the zip archive
            try
            {
                ZipFile.ExtractToDirectory(zipPath, extractFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting zip file: {ex.Message}");
                return;
            }

            // Supported Excel extensions
            string[] excelExtensions = { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

            // Process each extracted file
            foreach (string filePath in Directory.GetFiles(extractFolder, "*.*", SearchOption.AllDirectories))
            {
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (Array.IndexOf(excelExtensions, ext) < 0)
                    continue; // Skip non‑Excel files

                // Build the HTML output file name
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlFilePath = Path.Combine(htmlOutputFolder, htmlFileName);

                // Convert the Excel file to HTML using Aspose.Cells ConversionUtility
                try
                {
                    ConversionUtility.Convert(filePath, htmlFilePath);
                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML: {htmlFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            // Optional: clean up the temporary extraction folder
            try
            {
                Directory.Delete(extractFolder, true);
            }
            catch
            {
                // If deletion fails, ignore – the folder may be in use.
            }

            Console.WriteLine("All conversions completed.");
        }
    }
}
