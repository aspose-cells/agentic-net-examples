// Title: How to batch convert Excel workbooks (.xls, .xlsx, .xlsm) to HTML using Aspose.Cells in C#
// AI Prompts: Generate a C# console program that scans a directory, loads each .xls/.xlsx/.xlsm workbook with Aspose.Cells, and saves it as an .html file using the default HtmlSaveOptions. | Write code to iterate over all Excel files in a folder and perform bulk HTML conversion with Aspose.Cells while logging each successful conversion and any errors. | Create a C# script that handles exceptions and reports a summary of conversion results when exporting multiple Excel workbooks to HTML using Aspose.Cells.
// Common Searches: C# Aspose.Cells convert all Excel files in a folder to HTML | batch export of .xlsx and .xlsm workbooks to HTML with default settings using Aspose.Cells | how to use HtmlSaveOptions for bulk Excel to HTML conversion in a C# console app | Aspose.Cells directory iteration to generate HTML files from multiple workbooks
// Tags: Aspose.Cells bulk Excel to HTML conversion | C# HtmlSaveOptions default usage | process multiple .xls .xlsx .xlsm files | directory iteration with Aspose.Cells | error handling during Excel to HTML export

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that scans a specified folder, loads each .xls, .xlsx, or .xlsm workbook with Aspose.Cells, and saves it as an .html file using default HtmlSaveOptions, while logging successes and handling any errors.
class ExcelToHtmlBatchConverter
{
    static void Main(string[] args)
    {
        // Specify the folder containing Excel files.
        // You can change this path as needed or pass it via command line arguments.
        string sourceFolder = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

        // Verify that the folder exists.
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Folder not found: {sourceFolder}");
            return;
        }

        // Get all Excel files in the folder (including .xls, .xlsx, .xlsm).
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue; // Skip non‑Excel files.

            try
            {
                // Load the workbook from the Excel file.
                Workbook workbook = new Workbook(filePath);

                // Prepare the output HTML file path.
                string htmlFileName = Path.ChangeExtension(Path.GetFileName(filePath), ".html");
                string htmlFilePath = Path.Combine(sourceFolder, htmlFileName);

                // Save the workbook as HTML using default HtmlSaveOptions.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(); // defaults are sufficient.
                workbook.Save(htmlFilePath, saveOptions);

                Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {htmlFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
