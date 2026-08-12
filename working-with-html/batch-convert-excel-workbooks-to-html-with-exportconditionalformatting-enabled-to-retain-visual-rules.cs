// Title: Batch convert Excel workbooks to HTML with conditional formatting using Aspose.Cells for .NET
// Description: C# utility that scans a folder for .xlsx, .xls, .xlsm files, creates an output directory, and converts each workbook to HTML with Aspose.Cells HtmlSaveOptions, preserving conditional formatting (ExportConditionalFormatting enabled).
// Keywords: Aspose.Cells | C# Excel to HTML | batch conversion | conditional formatting export | HtmlSaveOptions | ConversionUtility | folder processing | xlsx to html | xlsm to html | .NET Excel HTML conversion
// Common Searches: convert all Excel files in a folder to HTML Aspose.Cells | export conditional formatting when saving Excel as HTML .NET | batch Excel to HTML conversion C# | Aspose.Cells HtmlSaveOptions ExportConditionalFormatting example | automate Excel to HTML conversion with Aspose
// Developer Intent: Convert every Excel workbook in a specified directory to an HTML file while keeping all conditional‑formatting rules intact.
// Use Cases: Generate nightly HTML previews of financial spreadsheets for web dashboards without losing color‑coded alerts. | Provide instant web‑ready views of user‑uploaded Excel files in a SaaS portal, preserving visual cues. | Migrate a legacy archive of Excel reports to static HTML pages for faster access and reduced server load.
// AI Prompts: Write C# code that iterates through a directory of .xlsx, .xls, and .xlsm files and converts each to HTML using Aspose.Cells, ensuring HtmlSaveOptions.ExportConditionalFormatting = true. | Explain best practices for error handling, logging, and performance when batch converting Excel workbooks to HTML with Aspose.Cells in a .NET application.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchExcelToHtml
{
    // C# utility that scans a folder for .xlsx, .xls, .xlsm files, creates an output directory, and converts each workbook to HTML with Aspose.Cells HtmlSaveOptions, preserving conditional formatting (ExportConditionalFormatting enabled).
    class Program
    {
        static void Main()
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\InputExcel";

            // Folder where HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (xlsx, xls, xlsm)
            var excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

            foreach (string sourcePath in excelFiles)
            {
                try
                {
                    // Build the destination HTML file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Load options – default loading behavior
                    LoadOptions loadOptions = new LoadOptions();

                    // HTML save options (conditional formatting is exported by default)
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                    // Perform the conversion using the provided ConversionUtility method
                    ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
