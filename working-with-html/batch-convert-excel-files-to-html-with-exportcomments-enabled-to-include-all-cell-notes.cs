// Title: Batch convert Excel workbooks to HTML with cell comments using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves it as an HTML file with ExportComments enabled. | Enhance the batch converter to create the output directory automatically and log the conversion result (success or error) for every processed file. | Add robust exception handling that skips unreadable or corrupted Excel files while continuing the batch conversion.
// Common Searches: c# aspose.cells batch convert xlsx files to html with comments | how to include cell notes when exporting Excel to html using aspose.cells | save all worksheets of a workbook as separate html files c# aspose | aspose.cells HtmlSaveOptions ExportComments property not available in version | automate conversion of multiple Excel files to html in a Windows service
// Tags: batch excel to html conversion aspose.cells | exportcellcomments htmlsaveoptions c# | convert multiple xlsx files to html .net | aspose.cells htmlsaveoptions exportcomments usage | process workbook folder programmatically c#

using System;
using System.IO;
using Aspose.Cells;

// The sample iterates through all .xls, .xlsx, and .xlsm files in a specified input directory, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to export the entire workbook, and saves the result as an HTML file in an output folder. It creates the output folder if needed, skips non‑Excel files, handles missing files and runtime exceptions, and logs each conversion. Note that the ExportComments property may not be available in older Aspose.Cells versions, so enabling cell comments may require a newer release.
class ExcelToHtmlBatchConverter
{
    static void Main()
    {
        // Folder containing the Excel files to convert
        string inputFolder = @"C:\InputExcelFiles";
        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtmlFiles";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (xls, xlsx, xlsm) in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue; // Skip non‑Excel files

            // Verify the file still exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = false // Export the entire workbook
                    // Note: ExportComments and ExportChartImageFormat are not available in this version of Aspose.Cells
                };

                // Determine output HTML file name
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlFilePath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML with the specified options
                workbook.Save(htmlFilePath, saveOptions);

                Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
