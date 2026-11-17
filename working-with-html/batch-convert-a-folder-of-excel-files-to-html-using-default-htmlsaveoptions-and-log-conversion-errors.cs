// Title: Batch convert Excel files in a folder to HTML with Aspose.Cells for .NET and log conversion errors
// Common Searches: C# code to batch convert .xls and .xlsx files to HTML using Aspose.Cells | How to export all Excel workbooks in a directory to HTML with error handling in .NET | Aspose.Cells example for converting a folder of Excel files to HTML and creating a log file | Program to save multiple Excel files as HTML using default HtmlSaveOptions in C# | Create a conversion log when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells batch conversion to HTML | C# convert Excel folder to HTML | HtmlSaveOptions default usage | Excel to HTML error handling | log conversion errors Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExcelToHtmlBatchConverter
{
    // The sample scans a source directory for Excel files, loads each workbook with Aspose.Cells, saves it as an HTML file using the default HtmlSaveOptions, and records any conversion exceptions to a log file in the destination folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Define source and destination folders
            string sourceFolder = @"C:\ExcelFiles";          // Folder containing Excel files
            string destinationFolder = @"C:\HtmlOutput";    // Folder where HTML files will be saved
            string logFilePath = Path.Combine(destinationFolder, "ConversionLog.txt");

            // Ensure the destination folder exists
            Directory.CreateDirectory(destinationFolder);

            // Prepare a list to collect error messages
            List<string> errorLog = new List<string>();

            // Get all Excel files in the source folder (including subfolders if needed)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                // Process only supported Excel extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                {
                    continue; // Skip non‑Excel files
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Prepare the output HTML file path (same name, .html extension)
                    string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                    string htmlFilePath = Path.Combine(destinationFolder, htmlFileName);

                    // Save the workbook as HTML using default HtmlSaveOptions
                    workbook.Save(htmlFilePath, SaveFormat.Html);
                }
                catch (Exception ex)
                {
                    // Record the error with file name and exception message
                    string errorMessage = $"Error converting '{Path.GetFileName(filePath)}': {ex.Message}";
                    errorLog.Add(errorMessage);
                }
            }

            // Write the error log to a text file (if any errors occurred)
            if (errorLog.Count > 0)
            {
                File.WriteAllLines(logFilePath, errorLog);
                Console.WriteLine($"Conversion completed with errors. See log at: {logFilePath}");
            }
            else
            {
                Console.WriteLine("All files converted successfully.");
            }
        }
    }
}
