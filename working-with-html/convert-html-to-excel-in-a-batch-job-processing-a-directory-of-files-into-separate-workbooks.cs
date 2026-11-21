// Title: Convert a folder of HTML files to individual XLSX workbooks using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells convert all HTML files in a directory to separate Excel files | How to batch process HTML to XLSX with Aspose.Cells .NET | Automate conversion of multiple .html documents to .xlsx using C# | Load HTML into Aspose.Cells Workbook and save as Excel in a loop | Directory scan and export each HTML file to Excel with Aspose.Cells
// Tags: Aspose.Cells batch HTML to XLSX conversion | C# directory iteration for file format conversion | Workbook.Load from HTML using Aspose.Cells | Workbook.Save as Xlsx programmatically | Automated HTML to Excel conversion .NET

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified input folder for *.html files, loads each file into an Aspose.Cells Workbook, and saves it as an .xlsx workbook with the same base name in an output folder. It creates the destination directory if it does not exist and logs any file‑specific or unexpected errors while continuing the batch process.
class HtmlToExcelBatchProcessor
{
    static void Main()
    {
        // Define source and destination directories
        string sourceDirectory = @"C:\InputHtmlFiles";
        string destinationDirectory = @"C:\OutputExcelFiles";

        try
        {
            // Verify source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
                return;
            }

            // Ensure the destination directory exists
            if (!Directory.Exists(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }

            // Process each HTML file in the source directory
            foreach (string htmlFilePath in Directory.GetFiles(sourceDirectory, "*.html"))
            {
                try
                {
                    // Verify the HTML file exists before loading
                    if (!File.Exists(htmlFilePath))
                    {
                        Console.WriteLine($"File not found: {htmlFilePath}");
                        continue;
                    }

                    // Load the HTML file into an Aspose.Cells Workbook
                    Workbook workbook = new Workbook(htmlFilePath);

                    // Build the output Excel file path (same name, .xlsx extension)
                    string excelFileName = Path.GetFileNameWithoutExtension(htmlFilePath) + ".xlsx";
                    string excelFilePath = Path.Combine(destinationDirectory, excelFileName);

                    // Save the workbook as an Excel file
                    workbook.Save(excelFilePath, SaveFormat.Xlsx);
                    Console.WriteLine($"Converted: {htmlFilePath} -> {excelFilePath}");
                }
                catch (Exception ex)
                {
                    // Log errors for individual files without stopping the batch
                    Console.WriteLine($"Error processing '{htmlFilePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
