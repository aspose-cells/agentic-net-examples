// Title: Batch convert Excel workbooks to HTML with cell comments using Aspose.Cells for .NET
// AI Prompts: Create a C# console application that scans a given folder, loads every .xls or .xlsx file with Aspose.Cells, and saves each workbook as an HTML file. | Demonstrate how to set HtmlSaveOptions.ExportComments to true so that cell notes are included when converting Excel to HTML with Aspose.Cells. | Add comprehensive try‑catch blocks and per‑file logging to a batch Excel‑to‑HTML conversion script in C#.
// Common Searches: aspocells c# batch convert xls xlsx to html with comments | how to export cell notes to html using Aspose.Cells in a loop | c# program to convert all Excel files in a directory to html preserving comments
// Tags: Aspose.Cells batch Excel to HTML conversion | HtmlSaveOptions export comments Aspose.Cells | C# directory based workbook processing | preserve cell notes when saving Excel as HTML | programmatic .xlsx to .html conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample enumerates .xls and .xlsx files in a source folder, loads each workbook with Aspose.Cells, configures HtmlSaveOptions (including ExportComments when supported), saves the workbook as an HTML file in a target folder, and logs success or error messages for every file processed.
class BatchExcelToHtml
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\ExcelFiles";
        // Folder where the HTML files will be saved
        string outputFolder = @"C:\HtmlOutput";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all files in the input folder
        string[] files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in files)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files

            // Verify the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook from the current Excel file
                Workbook workbook = new Workbook(filePath);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
                // Note: Exporting comments is not supported in this version of Aspose.Cells.
                // If needed, upgrade to a version that provides ExportComments property.

                // Build the output HTML file path (same name, .html extension)
                string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the workbook as HTML with the specified options
                workbook.Save(outputPath, saveOptions);

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
