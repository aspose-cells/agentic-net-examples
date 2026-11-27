// Title: PowerShell script that uses Aspose.Cells for .NET to batch convert Excel workbooks to HTML files
// Common Searches: powershell batch convert xlsx files to html using aspose.cells | how to call Aspose.Cells .NET library from PowerShell to export Excel as HTML | script to convert all Excel files in a folder to HTML with Aspose.Cells | automate Excel to HTML conversion for multiple spreadsheet formats via PowerShell | invoke Aspose.Cells SaveFormat.Html from PowerShell for bulk workbook conversion
// Tags: PowerShell invoke Aspose.Cells .NET for Excel to HTML conversion | batch conversion of Excel workbooks to HTML using SaveFormat.Html | process multiple spreadsheet extensions with Aspose.Cells in PowerShell | automated folder-level Excel to HTML export script | Aspose.Cells .NET library integration in PowerShell scripts

using System;
using System.IO;
using Aspose.Cells;

// A PowerShell script that validates input and output folders, scans a directory for supported Excel extensions, loads each workbook with Aspose.Cells, sets the workbook filename, saves it as an HTML file using SaveFormat.Html, disposes the workbook, and reports the number of successful and failed conversions.
class Program
{
    static void Main(string[] args)
    {
        // Validate arguments
        if (args.Length < 1 || args.Length > 2)
        {
            Console.WriteLine("Usage: Program.exe <InputFolderPath> [OutputFolderPath]");
            return;
        }

        string inputFolder = args[0];
        string outputFolder = args.Length == 2 ? args[1] : Path.Combine(inputFolder, "HtmlOutput");

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.Error.WriteLine($"Error: Input folder does not exist -> {inputFolder}");
            return;
        }

        // Ensure output folder exists
        try
        {
            Directory.CreateDirectory(outputFolder);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: Unable to create output folder '{outputFolder}'. {ex.Message}");
            return;
        }

        // Define supported Excel extensions
        string[] supportedExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods", ".xml" };

        // Get all files in the input folder (non‑recursive)
        string[] allFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);

        int processedCount = 0;
        int failedCount = 0;

        foreach (string filePath in allFiles)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (Array.IndexOf(supportedExtensions, ext) < 0)
            {
                // Skip non‑Excel files
                continue;
            }

            try
            {
                Console.WriteLine($"Processing: {Path.GetFileName(filePath)}");

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Set FileName property (helps with external references)
                workbook.FileName = filePath;

                // Determine output HTML file path
                string htmlFileName = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(filePath) + ".html");

                // Save as HTML
                workbook.Save(htmlFileName, SaveFormat.Html);

                // Clean up
                workbook.Dispose();

                Console.WriteLine($"Saved HTML: {htmlFileName}");
                processedCount++;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to convert '{filePath}': {ex.Message}");
                failedCount++;
            }
        }

        Console.WriteLine();
        Console.WriteLine($"Batch conversion completed. Success: {processedCount}, Failed: {failedCount}");
    }
}
