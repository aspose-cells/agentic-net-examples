// Title: C# batch conversion of Excel files to HTML with CSS custom properties for image deduplication (Aspose.Cells)
// Description: The example scans a directory, creates an output folder, and converts every .xls, .xlsx, or .xlsm workbook to an .html file. HtmlSaveOptions.EnableCssCustomProperties is enabled so identical images are referenced through a single CSS variable, cutting redundant image data.
// Keywords: Aspose.Cells | C# | Excel to HTML | batch conversion | EnableCssCustomProperties | image deduplication | HtmlSaveOptions | ConversionUtility | .NET automation | static web export
// Common Searches: Aspose.Cells convert all Excel files in a folder to HTML | EnableCssCustomProperties to remove duplicate images in HTML output | C# batch Excel to HTML conversion example | How to use HtmlSaveOptions for image deduplication | ConversionUtility batch processing Excel to HTML
// Developer Intent: Transform a collection of Excel workbooks into HTML pages while activating CSS custom properties to avoid repeated image files.
// Use Cases: Publish a suite of financial spreadsheets as lightweight web reports with minimal image overhead. | Automate the generation of static dashboard pages from Excel sources for intranet portals. | Archive legacy Excel documents as HTML for documentation sites, leveraging CSS variables to keep the payload small.
// AI Prompts: Generate a C# method that receives source and destination paths and uses Aspose.Cells to convert every Excel workbook to HTML with EnableCssCustomProperties turned on. | Explain the mechanism by which EnableCssCustomProperties reduces duplicate images in the HTML produced by Aspose.Cells and suggest ways to validate the result. | Outline robust error‑handling and logging strategies for a batch Excel‑to‑HTML conversion using ConversionUtility.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchExcelToHtml
{
    // The example scans a directory, creates an output folder, and converts every .xls, .xlsx, or .xlsm workbook to an .html file. HtmlSaveOptions.EnableCssCustomProperties is enabled so identical images are referenced through a single CSS variable, cutting redundant image data.
    class Program
    {
        static void Main()
        {
            // Folder containing Excel files to be converted
            string sourceFolder = @"C:\ExcelFiles";
            // Destination folder for generated HTML files
            string outputFolder = @"C:\HtmlOutput";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Get all Excel files (XLS, XLSX, XLSM) in the source folder
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                    continue; // Skip non‑Excel files

                // Prepare the HTML output file path (same name, .html extension)
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                try
                {
                    // Create HtmlSaveOptions and enable CSS custom properties for image deduplication
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        EnableCssCustomProperties = true
                    };

                    // Convert the Excel file to HTML with the specified options
                    ConversionUtility.Convert(
                        source: filePath,
                        loadOptions: null,          // No special load options required
                        saveAs: htmlPath,
                        saveOptions: htmlOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML with CSS custom properties.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
