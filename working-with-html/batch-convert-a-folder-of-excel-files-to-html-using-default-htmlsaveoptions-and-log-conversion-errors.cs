// Title: Batch convert a folder of Excel (.xls, .xlsx) files to HTML using Aspose.Cells for .NET and log conversion errors
// AI Prompts: Write a C# console program that scans a specified directory, loads each .xls or .xlsx workbook with Aspose.Cells, and saves it as an .html file using the default HtmlSaveOptions. | Modify the batch converter to capture any exceptions during the Save operation and append a detailed error entry to a log file in the output folder. | Extend the solution to recursively process subdirectories, preserving the folder hierarchy when creating corresponding HTML files.
// Common Searches: how to use Aspose.Cells to convert multiple Excel files to HTML in C# | C# batch conversion of .xls and .xlsx to .html with error logging | Aspose.Cells default HtmlSaveOptions example for folder processing | log conversion failures when saving Excel workbooks as HTML using Aspose.Cells | recursive Excel to HTML conversion using Aspose.Cells .NET
// Tags: Aspose.Cells batch Excel to HTML conversion | C# convert Excel workbooks to HTML | log Aspose.Cells conversion errors | default HtmlSaveOptions usage | process Excel files in directory with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates .xls and .xlsx files in a given source folder, loads each workbook with Aspose.Cells, saves it as an HTML file using the default HtmlSaveOptions, and writes any conversion exceptions to the error console or an optional log file.
class ExcelToHtmlBatchConverter
{
    static void Main(string[] args)
    {
        // Define the source folder containing Excel files
        string sourceFolder = @"C:\ExcelFiles";

        // Define the destination folder for HTML output
        string destinationFolder = @"C:\HtmlOutput";

        // Ensure the destination folder exists
        if (!Directory.Exists(destinationFolder))
        {
            Directory.CreateDirectory(destinationFolder);
        }

        // Get all Excel files (xls and xlsx) in the source folder
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files

            try
            {
                // Load the workbook from the Excel file
                Workbook workbook = new Workbook(filePath);

                // Prepare the output HTML file path (same name, .html extension)
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlFilePath = Path.Combine(destinationFolder, htmlFileName);

                // Save the workbook as HTML using default HtmlSaveOptions
                workbook.Save(htmlFilePath, SaveFormat.Html);
                
                Console.WriteLine($"Successfully converted: {Path.GetFileName(filePath)} -> {htmlFileName}");
            }
            catch (Exception ex)
            {
                // Log conversion errors
                Console.Error.WriteLine($"Error converting file '{filePath}': {ex.Message}");
                // Optionally, write to a log file
                // File.AppendAllText(Path.Combine(destinationFolder, "conversion_errors.log"),
                //     $"{DateTime.Now}: {filePath} - {ex}{Environment.NewLine}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
