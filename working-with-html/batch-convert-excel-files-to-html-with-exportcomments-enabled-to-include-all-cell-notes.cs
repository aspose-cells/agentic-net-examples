// Title: Batch convert Excel files to HTML with comments using Aspose.Cells for .NET
// Description: A C# utility that scans a given folder, loads each Excel workbook (xls, xlsx, xlsm, csv) and converts it to an HTML file. The conversion uses Aspose.Cells ConversionUtility with HtmlSaveOptions.IsExportComments enabled, ensuring all cell notes appear in the generated HTML.
// Keywords: Aspose.Cells | C# batch Excel to HTML | ExportComments | HtmlSaveOptions | ConversionUtility | convert multiple workbooks | cell notes to HTML | Excel to web conversion | CSV to HTML Aspose | automated spreadsheet conversion
// Common Searches: C# convert folder of Excel files to HTML with comments | Aspose.Cells batch export cell notes to HTML | How to use HtmlSaveOptions.IsExportComments | Convert xls xlsx csv to HTML programmatically | Aspose.Cells convert multiple workbooks to HTML
// Developer Intent: Automatically transform every Excel workbook in a directory into an HTML page while preserving all cell comments.
// Use Cases: Publish a collection of spreadsheets as web‑ready documentation that retains original reviewer notes. | Archive financial or audit reports in HTML format, keeping embedded comments for regulatory reference. | Provide instant HTML previews of uploaded CSV/XLSX files on a portal, showing cell notes alongside data.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert all Excel files in a directory to HTML with IsExportComments enabled and logs conversion failures. | Write a PowerShell script that calls the compiled .NET assembly to perform the same batch conversion, including error handling and progress output. | Explain how to extend the sample to also export cell formulas as HTML tooltips while still exporting comments.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# utility that scans a given folder, loads each Excel workbook (xls, xlsx, xlsm, csv) and converts it to an HTML file. The conversion uses Aspose.Cells ConversionUtility with HtmlSaveOptions.IsExportComments enabled, ensuring all cell notes appear in the generated HTML.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing source Excel files
        string inputFolder = @"C:\InputExcel";

        // Folder where HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files in the input folder (xls, xlsx, xlsm, csv)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string sourcePath in excelFiles)
        {
            string ext = Path.GetExtension(sourcePath).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".csv")
                continue; // Skip non‑Excel files

            // Verify the source file still exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"File not found, skipping: {sourcePath}");
                continue;
            }

            // Determine destination HTML file path
            string htmlFileName = Path.GetFileNameWithoutExtension(sourcePath) + ".html";
            string destPath = Path.Combine(outputFolder, htmlFileName);

            try
            {
                // Load options (auto‑detect format)
                LoadOptions loadOptions = new LoadOptions();

                // HTML save options with comments export enabled
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    IsExportComments = true
                };

                // Perform conversion using Aspose.Cells ConversionUtility
                ConversionUtility.Convert(sourcePath, loadOptions, destPath, htmlOptions);
                Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file '{sourcePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
