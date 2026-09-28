// Title: Batch convert Excel .xlsx workbooks to HTML with hidden worksheets included using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a folder for .xlsx files and saves each workbook as an HTML file, configuring Aspose.Cells to export hidden worksheets. | Demonstrate how to set HtmlSaveOptions.ExportHiddenWorksheet = true and ExportActiveWorksheetOnly = false for converting multiple Excel workbooks to HTML in a batch process.
// Common Searches: asp.net batch convert xlsx files to html including hidden sheets | c# Aspose.Cells export all worksheets to html | how to save hidden Excel worksheets as html with Aspose.Cells | process multiple Excel workbooks to html using Aspose.Cells SaveOptions | convert a folder of .xlsx files to html programmatically c#
// Tags: batch xlsx to html conversion Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet | export hidden worksheets to html c# | process multiple workbooks Aspose.Cells | convert Excel workbooks to html programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // The example scans a specified input directory for .xlsx files, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to include hidden worksheets and all sheets, and saves each workbook as an HTML file in an output directory, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files to convert
            string inputFolder = @"C:\InputExcelFiles";
            // Folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtmlFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: '{inputFolder}'.");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder (top level only)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: '{excelPath}'. Skipping.");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Configure HTML save options to export hidden worksheets
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        ExportHiddenWorksheet = true,          // Include hidden sheets in the output
                        ExportActiveWorksheetOnly = false      // Export all worksheets (including hidden)
                    };

                    // Determine output HTML file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
