// Title: Batch Convert All XLSX Files in a Folder to HTML Using Aspose.Cells (C#)
// Description: A C# console app that scans a specified directory, creates matching .html files in an output folder, and uses Aspose.Cells.Utility.ConversionUtility.Convert with default settings to transform each workbook to HTML. Includes folder validation and error handling.
// Keywords: Aspose.Cells batch conversion | C# XLSX to HTML | ConversionUtility Convert example | automate Excel to HTML | process multiple Excel files | default conversion settings | .NET Excel HTML export | folder based spreadsheet conversion
// Common Searches: convert all Excel files in a folder to HTML C# | Aspose.Cells batch convert XLSX to HTML example | C# code for converting multiple spreadsheets to web pages | default Aspose.Cells HTML export for many workbooks | how to automate Excel to HTML conversion with Aspose
// Developer Intent: Programmatically transform every .xlsx workbook in a given directory into an .html file using Aspose.Cells default conversion options.
// Use Cases: Publish a collection of financial models as static web pages for intranet access. | Generate web‑ready reports from daily Excel exports stored on a server. | Create archival HTML snapshots of spreadsheet data without manual steps.
// AI Prompts: Show a C# snippet that logs each conversion result to a CSV file while batch processing XLSX to HTML with Aspose.Cells. | Modify the code to recursively process subfolders and preserve the original folder hierarchy in the HTML output. | Explain how to apply custom CSS or embed styles in the generated HTML using Aspose.Cells conversion options.

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // A C# console app that scans a specified directory, creates matching .html files in an output folder, and uses Aspose.Cells.Utility.ConversionUtility.Convert with default settings to transform each workbook to HTML. Includes folder validation and error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the XLSX files.
            string sourceFolder = @"C:\InputXlsxFiles";

            // Folder where the HTML files will be saved.
            string outputFolder = @"C:\OutputHtmlFiles";

            // Verify source folder exists.
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Get all .xlsx files in the source folder.
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Build the output HTML file path.
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Convert the Excel file to HTML using default settings.
                    ConversionUtility.Convert(xlsxPath, htmlPath);

                    Console.WriteLine($"Converted: {xlsxPath} -> {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
