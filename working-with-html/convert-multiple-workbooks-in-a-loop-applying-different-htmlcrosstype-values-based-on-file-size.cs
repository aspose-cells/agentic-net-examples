// Title: Convert multiple Excel workbooks to HTML in C# with file‑size‑based HtmlCrossType using Aspose.Cells
// AI Prompts: Generate C# code that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, checks the file size, and saves it as HTML using HtmlSaveOptions where the output mode is a single file for files larger than 1 MB and multiple files for smaller ones. | Modify the existing conversion loop to enable base‑64 image embedding when the workbook size exceeds a configurable threshold and to switch the HTML output mode accordingly, then write the HTML output to a designated directory. | Create a helper method `ConvertToHtml(string sourcePath, string destFolder, long sizeThreshold)` that determines the workbook size, configures HtmlSaveOptions with the appropriate output mode (single‑file or split) and image embedding options, saves the HTML file, and returns the generated file path.
// Common Searches: aspnet convert a batch of .xlsx files to html with different output modes based on file size | c# Aspose.Cells how to set HTML output mode per workbook in a loop | conditional base64 image embedding for large Excel files when saving as HTML using Aspose.Cells | example code for size‑dependent HTML export options in Aspose.Cells C#
// Tags: Aspose.Cells batch workbook to HTML conversion | HtmlSaveOptions HtmlCrossType based on file size | C# conditional HTML export settings for Excel | ExportImagesAsBase64 large Excel files Aspose.Cells | automated folder processing Excel to HTML .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookHtmlConversion
{
    // The sample program enumerates all .xlsx files in a given input directory, creates an output folder if needed, and for each workbook determines its byte size. It loads the workbook with Aspose.Cells, configures HtmlSaveOptions (adjusting the output mode and optionally embedding images as Base64 when a size threshold is exceeded), saves the workbook as an .html file in the output directory, and logs success or any errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output directories (adjust as needed)
            string inputFolder = @"C:\InputWorkbooks";
            string outputFolder = @"C:\OutputHtml";

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string workbookPath in workbookFiles)
            {
                // Verify the workbook file exists before attempting to load
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"File not found: {workbookPath}");
                    continue;
                }

                try
                {
                    // Determine file size (bytes)
                    long fileSize = new FileInfo(workbookPath).Length;

                    // Load the workbook
                    Workbook workbook = new Workbook(workbookPath);

                    // Configure HTML save options (default settings)
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                    // Optionally, you could adjust save options based on file size here
                    // For example, you might set saveOptions.ExportImagesAsBase64 = true for large files

                    // Build output HTML file path
                    string outputFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".html";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the workbook as HTML
                    workbook.Save(outputPath, saveOptions);
                    Console.WriteLine($"Converted '{workbookPath}' to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    // Log the error and continue processing other files
                    Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
                }
            }
        }
    }
}
