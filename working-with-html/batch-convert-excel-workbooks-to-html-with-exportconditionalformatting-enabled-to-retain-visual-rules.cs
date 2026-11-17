// Title: Batch convert Excel workbooks to HTML with conditional formatting using Aspose.Cells for .NET (C#)
// Common Searches: C# code to convert all Excel files in a directory to HTML with Aspose.Cells preserving conditional formatting | How to batch export .xlsx workbooks to HTML using Aspose.Cells SaveOptions | Aspose.Cells example for exporting conditional formatting to HTML in C# | Programmatically convert multiple Excel worksheets to HTML files in .NET | Convert Excel files to HTML while keeping color scales and data bars with Aspose.Cells
// Tags: Aspose.Cells batch Excel to HTML conversion | HTMLSaveOptions C# Aspose.Cells | convert directory of .xlsx files to HTML Aspose | preserve conditional formatting in HTML export Aspose.Cells | C# iterate files and save workbooks as HTML

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified input folder for Excel workbooks (.xls, .xlsx, .xlsm), loads each file with Aspose.Cells, uses HtmlSaveOptions (which retains conditional formatting by default) to save the workbook as an HTML file in an output directory, creating the output folder if needed and logging each conversion.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing source Excel workbooks
            string inputFolder = @"C:\InputExcels";

            // Folder where the generated HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all files in the input folder (filter later for Excel extensions)
            string[] allFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);

            foreach (string filePath in allFiles)
            {
                // Verify the file has an Excel extension
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                    continue; // Skip non‑Excel files

                // Double‑check that the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Configure HTML save options (conditional formatting is exported by default)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Build the output HTML file path
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Converted '{filePath}' to '{htmlPath}'.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
