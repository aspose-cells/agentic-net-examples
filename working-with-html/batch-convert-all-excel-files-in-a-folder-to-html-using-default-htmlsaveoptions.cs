// Title: Convert all Excel files in a directory to HTML using Aspose.Cells for .NET (C#)
// Common Searches: C# code to convert every .xlsx file in a folder to HTML with Aspose.Cells | How to batch export Excel workbooks to HTML using Aspose.Cells .NET library | Aspose.Cells example for converting multiple Excel files to HTML in a loop | SaveFormat.Html batch conversion of Excel files in C# | Default HtmlSaveOptions usage for folder-wide Excel to HTML conversion
// Tags: Aspose.Cells batch Excel to HTML conversion | C# convert folder of workbooks to HTML | SaveFormat.Html default options | process multiple .xls .xlsx .xlsm files Aspose.Cells | directory-wide Excel to HTML conversion .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a specified source directory for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves it as an HTML file in an output folder using the default HtmlSaveOptions.
class ExcelToHtmlBatchConverter
{
    static void Main(string[] args)
    {
        // Specify the folder containing Excel files.
        string sourceFolder = @"C:\ExcelFiles";

        // Specify the folder where HTML files will be saved.
        string outputFolder = @"C:\HtmlOutput";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the source folder (including .xls, .xlsx, .xlsm).
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue; // Skip non‑Excel files.

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Prepare the output HTML file path.
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlFilePath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML using default HtmlSaveOptions.
                workbook.Save(htmlFilePath, SaveFormat.Html);
                
                Console.WriteLine($"Converted: {filePath} -> {htmlFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
