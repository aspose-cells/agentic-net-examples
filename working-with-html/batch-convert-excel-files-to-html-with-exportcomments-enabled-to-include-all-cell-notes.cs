// Title: Batch convert Excel (.xls, .xlsx) files to HTML with comments using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch convert folder of Excel workbooks to HTML preserving cell comments | How to export Excel cell notes to HTML for multiple files with Aspose.Cells .NET | Convert all .xls and .xlsx files in a directory to HTML including comments using Aspose.Cells | Aspose.Cells HtmlSaveOptions export comments when saving workbook as HTML in C# | Automate Excel to HTML conversion with comments for a collection of files in .NET
// Tags: Aspose.Cells batch conversion of workbooks | HTML export of Excel cell notes | HtmlSaveOptions SaveFormat.Html usage | C# convert xls and xlsx to HTML | process multiple Excel files programmatically

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a source directory for .xls and .xlsx files, loads each workbook with Aspose.Cells, applies HtmlSaveOptions (which export cell comments by default), and saves each workbook as an HTML file in a target folder, handling missing files and exceptions.
class ExcelToHtmlBatchConverter
{
    static void Main()
    {
        // Folder containing the Excel files to convert
        string sourceFolder = @"C:\ExcelFiles";
        // Destination folder for the generated HTML files
        string outputFolder = @"C:\HtmlOutput";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the source folder (including .xls and .xlsx)
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
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

                // Configure HTML save options (default settings already export comments)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Build the output HTML file path
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
