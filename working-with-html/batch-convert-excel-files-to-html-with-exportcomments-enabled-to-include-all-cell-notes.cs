// Title: Batch convert Excel (.xlsx) to HTML with comments using Aspose.Cells for .NET
// Description: A C# console utility that scans an "InputExcels" folder, creates an "OutputHtml" folder, and converts every .xlsx workbook to HTML. It uses Aspose.Cells ConversionUtility with HtmlSaveOptions.IsExportComments = true to retain all cell notes in the generated HTML files.
// Keywords: Aspose.Cells | C# batch Excel to HTML | ExportComments | HtmlSaveOptions | ConversionUtility | convert multiple xlsx to HTML | preserve cell comments | console utility | automate Excel to HTML | Aspose.Cells .NET
// Common Searches: batch convert excel files to html with comments aspose.cells | asp.net export cell notes to html for multiple workbooks | convert all xlsx in a folder to html using aspose.cells c# | htmlsaveoptions isexportcomments example | aspose.cells conversionutility batch processing
// Developer Intent: Automatically transform each .xlsx file in a directory into an HTML page while keeping every cell comment visible in the output.
// Use Cases: Generate web‑ready reports from a collection of spreadsheets, preserving analyst notes. | Add a step to CI/CD pipelines that publishes spreadsheet documentation as HTML with comments. | Deploy a lightweight console tool that watches a folder and converts newly added Excel files to HTML for immediate publishing.
// AI Prompts: Write a C# console app that monitors a folder and uses Aspose.Cells ConversionUtility to convert new .xlsx files to HTML with IsExportComments enabled. | Explain how HtmlSaveOptions.IsExportComments works when exporting Excel to HTML with Aspose.Cells. | Provide robust error‑handling patterns for batch converting Excel workbooks to HTML in a C# console application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# console utility that scans an "InputExcels" folder, creates an "OutputHtml" folder, and converts every .xlsx workbook to HTML. It uses Aspose.Cells ConversionUtility with HtmlSaveOptions.IsExportComments = true to retain all cell notes in the generated HTML files.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = "InputExcels";

        // Ensure the input directory exists; create it if missing
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder \"{inputFolder}\" does not exist. Creating it.");
            Directory.CreateDirectory(inputFolder);
            Console.WriteLine("Place Excel files in the input folder and rerun the program.");
            return;
        }

        // Folder where the HTML files will be saved
        string outputFolder = "OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (you can add more extensions if needed)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        if (excelFiles.Length == 0)
        {
            Console.WriteLine("No Excel files found in the input folder.");
            return;
        }

        foreach (string sourcePath in excelFiles)
        {
            // Verify the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                continue;
            }

            // Build the destination HTML file path
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

            try
            {
                // Load options – default settings are sufficient for most cases
                LoadOptions loadOptions = new LoadOptions();

                // HTML save options with ExportComments enabled to include cell notes
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    IsExportComments = true
                };

                // Perform the conversion using Aspose.Cells ConversionUtility
                ConversionUtility.Convert(sourcePath, loadOptions, destPath, htmlOptions);

                Console.WriteLine($"Converted '{sourcePath}' to '{destPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
            }
        }
    }
}
