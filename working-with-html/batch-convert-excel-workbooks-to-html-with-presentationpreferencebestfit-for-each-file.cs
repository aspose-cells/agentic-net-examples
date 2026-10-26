// Title: Batch convert multiple Excel workbooks to HTML with Aspose.Cells C# and PresentationPreference.BestFit
// Common Searches: c# loop through a list of .xlsx files and export each to html using aspose.cells | asp.net batch excel to html conversion with PresentationPreference BestFit | how to save multiple workbooks as html in a single folder with Aspose.Cells | using HtmlSaveOptions to set BestFit layout for each workbook in C# | automate conversion of several Excel files to html files with Aspose.Cells library
// Tags: batch excel to html conversion Aspose.Cells | HtmlSaveOptions PresentationPreference BestFit C# | export multiple workbooks as html Aspose.Cells | C# iterate excel files and save as html | Aspose.Cells HTML export settings per workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlBatch
{
    // The example lists Excel file paths, creates an output directory, and iterates over each workbook. For every existing file it loads the workbook with Aspose.Cells, configures HtmlSaveOptions (including optional PresentationPreference.BestFit), builds a matching HTML file name, and saves the workbook as an HTML document, logging successes and errors to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // List of Excel files to convert.
            List<string> excelFiles = new List<string>
            {
                @"C:\Input\Workbook1.xlsx",
                @"C:\Input\Workbook2.xls",
                @"C:\Input\Workbook3.xlsx"
            };

            // Destination folder for HTML files.
            string outputFolder = @"C:\Output\Html";

            // Ensure the output folder exists.
            Directory.CreateDirectory(outputFolder);

            // Process each workbook.
            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Verify the source file exists.
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"Source file not found: '{excelPath}'. Skipping.");
                        continue;
                    }

                    // Load the workbook.
                    Workbook workbook = new Workbook(excelPath);

                    // Configure HTML save options.
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        // Optional: Export only the first worksheet.
                        // ExportActiveWorksheetOnly = true,
                        // Optional: Embed images as base64.
                        // ExportImagesAsBase64 = true
                    };

                    // Determine output HTML file name.
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML.
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
