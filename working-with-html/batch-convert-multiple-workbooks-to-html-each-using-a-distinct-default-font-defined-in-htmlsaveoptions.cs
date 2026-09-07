// Title: Batch convert multiple Excel workbooks to HTML with a custom default font for each file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over a collection of Excel file paths, assigns a specific HtmlSaveOptions.DefaultFontName for each workbook, and saves the output as HTML with Aspose.Cells. | Write a .NET routine that checks whether each source Excel file exists, creates the target HTML folder if it doesn't exist, and exports the workbook to HTML using a distinct default font defined in HtmlSaveOptions. | Add robust exception handling to a batch HTML export loop that logs missing source files and continues processing the remaining workbooks.
// Common Searches: Aspose.Cells C# batch export Excel to HTML with different default fonts per workbook | How to set HtmlSaveOptions.DefaultFontName for each file in a bulk conversion loop | C# code sample for converting multiple .xlsx files to HTML using Aspose.Cells and custom fonts | Save Excel workbook as HTML with a specific font using Aspose.Cells .NET API
// Tags: batch Excel to HTML conversion Aspose.Cells | HtmlSaveOptions.DefaultFontName per workbook | C# iterate workbook list export HTML | create output directory before saving Aspose.Cells | exception handling batch HTML export Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchHtmlExport
{
    // The example defines a list of (sourcePath, htmlPath, defaultFont) tuples, validates each source file, ensures the output directory exists, loads the workbook, configures HtmlSaveOptions.DefaultFontName with the specified font, and saves the workbook as HTML. Errors are caught and logged, allowing the batch process to continue until all workbooks are processed.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the batch conversion list: each entry contains the source workbook path,
            // the target HTML path, and the default font to be used during HTML export.
            var conversionTasks = new List<(string sourcePath, string htmlPath, string defaultFont)>
            {
                (@"C:\Workbooks\Report1.xlsx", @"C:\HtmlOutputs\Report1.html", "Arial"),
                (@"C:\Workbooks\Report2.xlsx", @"C:\HtmlOutputs\Report2.html", "Times New Roman"),
                (@"C:\Workbooks\Report3.xlsx", @"C:\HtmlOutputs\Report3.html", "Calibri")
                // Add more entries as needed
            };

            // Process each workbook in the batch
            foreach (var task in conversionTasks)
            {
                try
                {
                    // Verify source file exists
                    if (!File.Exists(task.sourcePath))
                    {
                        Console.WriteLine($"Source file not found: {task.sourcePath}");
                        continue;
                    }

                    // Ensure output directory exists
                    string outputDir = Path.GetDirectoryName(task.htmlPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Load the workbook from the specified file
                    Workbook workbook = new Workbook(task.sourcePath);

                    // Configure HTML save options with the distinct default font
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                    htmlOptions.DefaultFontName = task.defaultFont; // Set the default font for this export

                    // Save the workbook as an HTML file using the configured options
                    workbook.Save(task.htmlPath, htmlOptions);

                    Console.WriteLine($"Converted '{task.sourcePath}' to HTML with default font '{task.defaultFont}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{task.sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
