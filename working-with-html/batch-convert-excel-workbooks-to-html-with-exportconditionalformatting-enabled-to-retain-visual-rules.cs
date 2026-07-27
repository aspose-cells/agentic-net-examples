// Title: Batch Convert Excel Workbooks to HTML with Conditional Formatting Using Aspose.Cells for .NET
// Description: A C# console app that scans a given folder, creates an output directory if needed, and converts every XLSX, XLSM, or XLS file to HTML. It uses Aspose.Cells LoadOptions and HtmlSaveOptions with ExportActiveWorksheetOnly = false, ExportWorkbookProperties = true, ExportWorksheetProperties = true, and ExportConditionalFormatting = true to retain visual rules and metadata. ConversionUtility handles the transformation while the code logs successes and errors.
// Keywords: Aspose.Cells batch Excel to HTML | C# export conditional formatting to HTML | HtmlSaveOptions ExportConditionalFormatting | convert folder of Excel files Aspose.Cells | retain workbook properties HTML conversion | Aspose.Cells ConversionUtility example | bulk Excel to HTML C#
// Common Searches: How to batch convert Excel files to HTML with conditional formatting using Aspose.Cells | C# Aspose.Cells export whole workbook to HTML preserving formatting | Convert multiple .xlsx files to HTML programmatically | Aspose.Cells HtmlSaveOptions ExportConditionalFormatting example | Automate Excel to HTML conversion in .NET
// Developer Intent: Convert all Excel workbooks in a directory to HTML files while preserving workbook/worksheet properties and conditional formatting rules.
// Use Cases: Publish a collection of financial spreadsheets as web‑ready HTML reports that keep color‑coded rules. | Archive engineering calculation sheets with their original conditional formatting for documentation portals. | Process user‑uploaded Excel files on a server, generating HTML previews that display the same visual cues as the source workbooks.
// AI Prompts: Update the code to set HtmlSaveOptions.ExportConditionalFormatting = true and explain its effect. | Add a CSV logger that records source file, destination HTML path, and conversion status for each workbook. | Show how to limit the conversion to specific worksheets while still using ConversionUtility. | Optimize the batch process for large numbers of files by using parallel execution with Aspose.Cells. | Explain how to customize the generated HTML (styles, embedded images) using HtmlSaveOptions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchExcelToHtml
{
    // A C# console app that scans a given folder, creates an output directory if needed, and converts every XLSX, XLSM, or XLS file to HTML. It uses Aspose.Cells LoadOptions and HtmlSaveOptions with ExportActiveWorksheetOnly = false, ExportWorkbookProperties = true, ExportWorksheetProperties = true, and ExportConditionalFormatting = true to retain visual rules and metadata. ConversionUtility handles the transformation while the code logs successes and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel workbooks to be converted
            string inputFolder = @"C:\InputExcelFiles";

            // Folder where the resulting HTML files will be saved
            string outputFolder = @"C:\OutputHtmlFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files (XLSX, XLSM, XLS) in the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string excelPath in excelFiles)
            {
                // Filter supported Excel formats
                string ext = Path.GetExtension(excelPath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xlsm" && ext != ".xls")
                    continue;

                // Verify the file actually exists (prevents FileNotFoundException)
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    continue;
                }

                // Prepare the destination HTML file path
                string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                try
                {
                    // Load options – let Aspose.Cells detect the format automatically
                    LoadOptions loadOptions = new LoadOptions();

                    // HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        ExportActiveWorksheetOnly = false,    // export the whole workbook
                        ExportWorkbookProperties = true,      // keep workbook properties
                        ExportWorksheetProperties = true      // keep worksheet properties
                    };

                    // Perform the conversion using the provided ConversionUtility method
                    ConversionUtility.Convert(excelPath, loadOptions, htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
