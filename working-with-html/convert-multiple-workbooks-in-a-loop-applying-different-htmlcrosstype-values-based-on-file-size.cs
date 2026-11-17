// Title: Batch convert Excel workbooks to HTML with file‑size‑based HtmlSaveOptions in C# (Aspose.Cells)
// Common Searches: C# code to convert all .xlsx files in a folder to HTML using Aspose.Cells | how to set HtmlSaveOptions.ExportImagesAsBase64 based on workbook size | loop through multiple workbooks and apply different HtmlSaveOptions per file | save large Excel files as HTML with external images using Aspose.Cells | automate Excel to HTML conversion with size thresholds in .NET
// Tags: batch Excel to HTML conversion Aspose.Cells C# | size‑dependent HtmlSaveOptions configuration | ExportImagesAsBase64 conditional setting | process multiple workbooks in a directory | HTML export for large Excel files Aspose.Cells | file size based HTML save options .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a directory for *.xlsx files, checks each file’s byte size, configures Aspose.Cells HtmlSaveOptions (toggling ExportImagesAsBase64) according to defined size thresholds, loads the workbook, and saves it as an HTML file in a target folder, with error handling for missing files.
class WorkbookHtmlConverter
{
    static void Main()
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string workbookPath in workbookFiles)
        {
            // Verify the workbook file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                continue;
            }

            try
            {
                // Determine the file size (in bytes)
                long fileSize = new FileInfo(workbookPath).Length;

                // Create HtmlSaveOptions (cross‑type handling is not available in this version)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Optionally adjust other settings based on size (example: export images as base64 for smaller files)
                if (fileSize < 1_000_000) // less than 1 MB
                {
                    htmlOptions.ExportImagesAsBase64 = true;
                }
                else if (fileSize < 5_000_000) // between 1 MB and 5 MB
                {
                    htmlOptions.ExportImagesAsBase64 = false;
                }
                else // 5 MB or larger
                {
                    // For large files, keep images as separate files to reduce HTML size
                    htmlOptions.ExportImagesAsBase64 = false;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Build the output HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(workbookPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML with the configured options
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Converted '{workbookPath}' to '{htmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
            }
        }
    }
}
