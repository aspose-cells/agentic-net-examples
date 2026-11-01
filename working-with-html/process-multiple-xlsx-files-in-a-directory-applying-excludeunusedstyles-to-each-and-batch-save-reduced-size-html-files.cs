// Title: Batch convert multiple XLSX workbooks to minimized HTML using Aspose.Cells RemoveUnusedStyles in C#
// AI Prompts: Write a C# console application that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, invokes Workbook.RemoveUnusedStyles, and saves the result as HTML using HtmlSaveOptions tuned for small output. | Generate C# code that processes every Excel file in a directory, disables Base64 image embedding and hidden worksheet export, and writes reduced‑size HTML files to a target folder.
// Common Searches: c# batch convert xlsx files to html with aspose.cells removeunusedstyles | how to reduce html size when exporting Excel using Aspose.Cells | aspnet console app to convert a folder of Excel workbooks to html | aspose.cells HtmlSaveOptions ExportImagesAsBase64 false example | process multiple Excel files and save as html using Aspose.Cells C#
// Tags: batch xlsx to html conversion aspose.cells | removeunusedstyles workbook c# | htmlsaveoptions exportimagesasbase64 false | exporthiddenworksheet false aspose.cells | reduce html output size from excel c#

using System;
using System.IO;
using Aspose.Cells;

// The program iterates over all .xlsx files in a specified input directory, loads each workbook with Aspose.Cells, removes unused styles to shrink the workbook, configures HtmlSaveOptions to avoid Base64‑encoded images and hidden worksheets, and saves each workbook as a minimized HTML file with the same name into an output directory.
class Program
{
    static void Main()
    {
        try
        {
            // Define the directory containing the source XLSX files
            string inputDirectory = @"C:\InputXlsxFiles";

            // Define the directory where the reduced‑size HTML files will be saved
            string outputDirectory = @"C:\OutputHtmlFiles";

            // Verify input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Get all XLSX files in the input directory
            string[] xlsxFiles = Directory.GetFiles(inputDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            // Process each workbook
            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Load the workbook from the XLSX file
                    Workbook workbook = new Workbook(xlsxPath);

                    // Remove any styles that are not used in the workbook to reduce size
                    workbook.RemoveUnusedStyles();

                    // Configure HTML save options for a smaller output
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        // Do not embed images as Base64 strings (creates separate image files)
                        ExportImagesAsBase64 = false,

                        // Do not export hidden worksheets
                        ExportHiddenWorksheet = false,

                        // Export all worksheets (set to true if you want only the active one)
                        ExportActiveWorksheetOnly = false
                    };

                    // Build the output HTML file path (same name as the source file)
                    string htmlFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".html";
                    string htmlPath = Path.Combine(outputDirectory, htmlFileName);

                    // Save the workbook as an HTML file using the configured options
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {htmlFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed. HTML files saved to: " + outputDirectory);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
