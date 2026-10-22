// Title: Batch convert multiple Excel workbooks to HTML with individual default fonts using Aspose.Cells for .NET
// AI Prompts: Generate a C# program that iterates over a collection of Excel file paths, loads each workbook with Aspose.Cells, and saves it as HTML using HtmlSaveOptions where the DefaultFontName is assigned uniquely for each workbook. | Write a C# script that validates the existence of source files, creates missing output directories, and exports each workbook to HTML with grid lines, embedded base64 images, and a custom default font specified per conversion.
// Common Searches: asp.net batch export Excel to HTML with different default fonts per file | c# Aspose.Cells convert multiple .xlsx to .html using HtmlSaveOptions DefaultFontName | how to set a unique default font for each HTML export in Aspose.Cells | automate Excel to HTML conversion with gridlines and base64 images in C#
// Tags: Aspose.Cells multiple workbook HTML export | set specific font for HTML export in Aspose.Cells | C# convert Excel to HTML with embedded images | base64 image embedding in Aspose.Cells HTML | automated Excel to HTML conversion script .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example defines a list of conversion tasks, each containing a source Excel file, a destination HTML file, and a default font name. For every task it checks the source file, creates the output folder if necessary, loads the workbook, configures HtmlSaveOptions to export grid lines, embed images as base64, and apply the specified DefaultFontName, then saves the workbook as HTML and logs the outcome, handling any errors and reporting completion.
class BatchHtmlConverter
{
    static void Main()
    {
        // List of conversion tasks: source workbook, destination HTML, and default font for HTML export.
        var conversionTasks = new List<(string sourcePath, string htmlPath, string defaultFont)>
        {
            (@"C:\Workbooks\Report1.xlsx", @"C:\HtmlOutputs\Report1.html", "Arial"),
            (@"C:\Workbooks\Report2.xlsx", @"C:\HtmlOutputs\Report2.html", "Times New Roman"),
            (@"C:\Workbooks\Report3.xlsx", @"C:\HtmlOutputs\Report3.html", "Calibri")
            // Add more entries as needed.
        };

        foreach (var task in conversionTasks)
        {
            try
            {
                // Verify source file exists.
                if (!File.Exists(task.sourcePath))
                {
                    Console.WriteLine($"Source file not found: {task.sourcePath}");
                    continue;
                }

                // Ensure the output directory exists.
                string? outputDir = Path.GetDirectoryName(task.htmlPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(task.sourcePath);

                // Configure HTML save options, including the default font.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,
                    ExportImagesAsBase64 = true,
                    DefaultFontName = task.defaultFont // Correct property name
                };

                // Save as HTML.
                workbook.Save(task.htmlPath, saveOptions);

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
