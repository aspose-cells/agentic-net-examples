// Title: Convert multiple Excel workbooks to HTML with a unique default font for each workbook using Aspose.Cells for .NET
// Common Searches: C# batch convert Excel files to HTML with different default fonts using Aspose.Cells | set default font per workbook before saving as HTML in Aspose.Cells .NET | how to apply a specific default font when exporting Excel to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions custom default font for each workbook in a batch | automate conversion of several .xlsx files to .html with individual font settings in C#
// Tags: batch workbook to HTML conversion Aspose.Cells | per-workbook default font HtmlSaveOptions | set workbook default style font C# | export Excel to HTML with custom default font | Aspose.Cells embed images base64 HTML export

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example defines a list of conversion tasks, each containing a source .xlsx path, a destination .html path, and a default font. It iterates through the list, validates files, creates output folders, loads each workbook, assigns the task's font to workbook.DefaultStyle.Font.Name, configures HtmlSaveOptions (including Base64 image embedding), and saves the workbook as HTML, handling errors and reporting progress.
class BatchHtmlConverter
{
    // Represents a conversion task: source workbook, destination HTML file, and the default font to use.
    public class ConversionTask
    {
        public string SourcePath { get; set; }
        public string DestinationPath { get; set; }
        public string DefaultFont { get; set; }

        public ConversionTask(string sourcePath, string destinationPath, string defaultFont)
        {
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
            DefaultFont = defaultFont;
        }
    }

    static void Main()
    {
        // Define the batch of workbooks to convert, each with its own default font.
        var tasks = new List<ConversionTask>
        {
            new ConversionTask(@"C:\Input\Report1.xlsx", @"C:\Output\Report1.html", "Arial"),
            new ConversionTask(@"C:\Input\Report2.xlsx", @"C:\Output\Report2.html", "Times New Roman"),
            new ConversionTask(@"C:\Input\Report3.xlsx", @"C:\Output\Report3.html", "Calibri")
            // Add more tasks as needed.
        };

        foreach (var task in tasks)
        {
            try
            {
                // Verify source file exists.
                if (!File.Exists(task.SourcePath))
                {
                    Console.WriteLine($"Source file not found: {task.SourcePath}");
                    continue;
                }

                // Ensure output directory exists.
                var outputDir = Path.GetDirectoryName(task.DestinationPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook.
                var workbook = new Workbook(task.SourcePath);

                // Set the default font for the workbook (used when a cell has no explicit font).
                // Aspose.Cells uses the DefaultStyle's Font.Name to define the default font.
                workbook.DefaultStyle.Font.Name = task.DefaultFont;

                // Configure HTML save options.
                var saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true,          // Embed images as Base64.
                    ExportActiveWorksheetOnly = false    // Export all worksheets.
                };

                // Save as HTML.
                workbook.Save(task.DestinationPath, saveOptions);
                Console.WriteLine($"Converted: {task.SourcePath} -> {task.DestinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {task.SourcePath}: {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
