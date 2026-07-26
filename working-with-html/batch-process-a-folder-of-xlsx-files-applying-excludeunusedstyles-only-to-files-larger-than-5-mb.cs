// Title: Batch Convert XLSX to HTML with Size‑Based ExcludeUnusedStyles Using Aspose.Cells for .NET
// Description: A C# utility that scans a folder for *.xlsx files, checks each file's size, and exports it to HTML. When a workbook exceeds 5 MB, HtmlSaveOptions.ExcludeUnusedStyles is enabled to omit unused CSS, reducing the HTML payload. The script creates the output directory, handles per‑file errors, and disposes resources correctly.
// Keywords: Aspose.Cells | C# batch Excel to HTML | ExcludeUnusedStyles | conditional HtmlSaveOptions | large workbook export | file size based conversion | HTML export from XLSX | automated Excel processing
// Common Searches: How to export multiple XLSX files to HTML with Aspose.Cells | Set ExcludeUnusedStyles only for large Excel files in C# | Batch convert Excel to HTML based on file size | Aspose.Cells conditional HtmlSaveOptions example | C# script to process folder of Excel workbooks
// Developer Intent: Automatically convert every XLSX in a directory to HTML, turning on ExcludeUnusedStyles only for files larger than 5 MB.
// Use Cases: Shrink HTML output for heavy reports while keeping full styling for small workbooks. | Schedule nightly conversion of a reports folder to web‑ready HTML with size‑aware optimization. | Integrate into a document pipeline that selectively removes unused CSS from large Excel exports.
// AI Prompts: Write C# code that iterates through a directory of .xlsx files and saves each as HTML with Aspose.Cells, enabling HtmlSaveOptions.ExcludeUnusedStyles only when the file size exceeds 5 MB. | Add logging to the batch conversion script so that errors are written to a log file while the process continues with remaining files. | Extend the program to copy the generated HTML files to a backup folder after each successful save.

using System;
using System.IO;
using Aspose.Cells;

// A C# utility that scans a folder for *.xlsx files, checks each file's size, and exports it to HTML. When a workbook exceeds 5 MB, HtmlSaveOptions.ExcludeUnusedStyles is enabled to omit unused CSS, reducing the HTML payload. The script creates the output directory, handles per‑file errors, and disposes resources correctly.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing the source XLSX files
            string sourceFolder = @"C:\InputFolder";

            // Folder where the processed HTML files will be saved
            string outputFolder = @"C:\OutputFolder";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process each XLSX file in the source folder
            foreach (string xlsxPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Determine file size
                    FileInfo fileInfo = new FileInfo(xlsxPath);
                    bool isLarge = fileInfo.Length > 5 * 1024 * 1024; // larger than 5 MB

                    // Load the workbook within a using block for proper disposal
                    using (Workbook workbook = new Workbook(xlsxPath))
                    {
                        // Prepare HTML save options and set ExcludeUnusedStyles based on size
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            ExcludeUnusedStyles = isLarge // true for large files, false otherwise
                        };

                        // Build output HTML file path
                        string htmlPath = Path.Combine(outputFolder,
                            Path.GetFileNameWithoutExtension(xlsxPath) + ".html");

                        // Save the workbook as HTML with the configured options
                        workbook.Save(htmlPath, htmlOptions);
                        Console.WriteLine($"Saved HTML: {htmlPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
