// Title: Batch Convert Excel (.xlsx) Files to HTML with Gridlines using Aspose.Cells for .NET
// Description: A C# console utility that scans a folder for .xlsx workbooks, creates an output directory, and converts each file to HTML with visible gridlines by configuring HtmlSaveOptions.ExportGridLines and invoking ConversionUtility. Includes basic error handling and logging.
// Keywords: Aspose.Cells | C# batch Excel to HTML | ExportGridLines | HtmlSaveOptions | ConversionUtility | .NET Excel conversion | HTML preview of Excel | bulk Excel to HTML | gridlines in HTML | Aspose.Cells example
// Common Searches: Aspose.Cells batch convert Excel to HTML | C# export Excel gridlines to HTML | Convert all .xlsx files in folder to HTML Aspose | HtmlSaveOptions ExportGridLines example | How to use ConversionUtility for Excel to HTML | Bulk Excel to HTML conversion .NET
// Developer Intent: Automatically transform every .xlsx workbook in a specified directory into an HTML file that displays the original spreadsheet’s gridlines.
// Use Cases: Generate web‑ready previews of a library of Excel reports while preserving cell borders. | Automate nightly publishing of financial worksheets to an intranet by converting them to HTML with gridlines. | Create a batch processing job that converts newly uploaded Excel files to browser‑friendly HTML for quick viewing.
// AI Prompts: Write C# code that iterates through a folder of .xlsx files and converts each to HTML with ExportGridLines enabled using Aspose.Cells. | Show how to modify the batch conversion program to export only the active worksheet instead of the whole workbook. | Suggest robust error‑handling patterns for bulk Excel‑to‑HTML conversion with Aspose.Cells in a .NET console app.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# console utility that scans a folder for .xlsx workbooks, creates an output directory, and converts each file to HTML with visible gridlines by configuring HtmlSaveOptions.ExportGridLines and invoking ConversionUtility. Includes basic error handling and logging.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing source Excel files
        string inputFolder = @"InputExcels";
        // Folder where HTML files will be saved
        string outputFolder = @"OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify that the input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder '{inputFolder}' does not exist. No files to process.");
            return;
        }

        try
        {
            // Process each .xlsx file in the input folder
            foreach (string sourcePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Build the destination HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Load options for the source Excel file (optional, can be default)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Configure HTML save options with gridlines enabled
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true // Show spreadsheet gridlines in HTML
                    // ExportActiveWorksheetOnly = true // Uncomment to export only the active sheet
                };

                try
                {
                    // Perform the conversion using the provided utility method
                    ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);
                    Console.WriteLine($"Converted '{sourcePath}' to '{destPath}' with gridlines.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
