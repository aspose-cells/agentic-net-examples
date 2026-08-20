// Title: C# – Batch Convert XLSX Files >5 MB to HTML with Aspose.Cells ExcludeUnusedStyles
// Description: A C# console application that scans a folder, loads each .xlsx larger than 5 MB with Aspose.Cells, sets HtmlSaveOptions.ExcludeUnusedStyles = true, and saves the workbook as an HTML file in a target directory while logging successes and errors.
// Keywords: Aspose.Cells | C# Excel to HTML conversion | ExcludeUnusedStyles | HtmlSaveOptions | batch convert XLSX to HTML | process large Excel files | convert files over 5 MB | .NET Excel HTML export | folder processing C# | Aspose.Cells example
// Common Searches: convert large xlsx to html aspose.cells | c# batch excel to html exclude unused styles | htmlsaveoptions excludeunusedstyles example | process folder of xlsx files .net | asp.net convert excel files larger than 5mb to html
// Developer Intent: Convert only XLSX workbooks larger than 5 MB in a directory to HTML using Aspose.Cells with unused styles excluded.
// Use Cases: Generate compact HTML reports from big Excel workbooks for web portals. | Automate nightly conversion of financial spreadsheets that exceed a size threshold for intranet dashboards. | Archive large Excel files as lightweight HTML pages in a document management system. | Create HTML previews for user uploads, processing only files that meet a minimum size to reduce server load.
// AI Prompts: Provide a C# function that iterates through a directory and converts every .xlsx file larger than 5 MB to HTML using Aspose.Cells with ExcludeUnusedStyles enabled. | Show how to add CSV logging of source path, destination path, file size, and conversion status to the batch program. | Explain how to extend the code to support .xls files and to apply a custom CSS stylesheet via HtmlSaveOptions. | Suggest ways to parallelize the conversion for faster processing on multi‑core machines. | Demonstrate how to wrap the logic into a reusable class library with dependency injection.

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that scans a folder, loads each .xlsx larger than 5 MB with Aspose.Cells, sets HtmlSaveOptions.ExcludeUnusedStyles = true, and saves the workbook as an HTML file in a target directory while logging successes and errors.
class Program
{
    static void Main()
    {
        // Folder containing the source XLSX files
        string sourceFolder = @"C:\InputFolder";

        // Folder where the processed HTML files will be saved
        string outputFolder = @"C:\OutputFolder";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Process each .xlsx file in the source folder
        foreach (string xlsxPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            try
            {
                // Verify the file still exists before processing
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found: {xlsxPath}");
                    continue;
                }

                FileInfo fileInfo = new FileInfo(xlsxPath);

                // Apply ExcludeUnusedStyles only if the file size exceeds 5 MB
                if (fileInfo.Length > 5 * 1024 * 1024)
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(xlsxPath);

                    // Create HTML save options and enable exclusion of unused styles
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExcludeUnusedStyles = true
                    };

                    // Build the output HTML file path
                    string htmlPath = Path.Combine(
                        outputFolder,
                        Path.GetFileNameWithoutExtension(xlsxPath) + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, htmlOptions);
                    Console.WriteLine($"Converted: {xlsxPath} -> {htmlPath}");
                }
            }
            catch (Exception ex)
            {
                // Log any errors for the current file and continue processing others
                Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
            }
        }
    }
}
