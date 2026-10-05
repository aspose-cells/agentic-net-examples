// Title: Batch convert a folder of XLSX workbooks to HTML with Aspose.Cells while excluding document properties
// AI Prompts: Iterate over every .xlsx file in a source directory, use Aspose.Cells to save each workbook as HTML, and set HtmlSaveOptions.ExportDocumentProperties to false before writing the files to a target folder. | Create a C# utility that loads each Excel file from a given path, converts all worksheets to HTML, disables metadata export, and stores the resulting .html files in a separate output directory.
// Common Searches: how to batch convert xlsx files to html using aspose.cells c# | c# export excel workbook to html without document properties | aspose.cells hide metadata when saving excel as html | convert all worksheets in multiple Excel files to html folder c# | privacy compliant html export from excel workbooks asp.net
// Tags: batch xlsx to html conversion Aspose.Cells | HtmlSaveOptions ExportDocumentProperties false | exclude workbook metadata in html export | convert folder of Excel files to html c# | export all worksheets to html Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsxToHtml
{
    // The sample iterates through all .xlsx files in a specified input folder, loads each workbook with Aspose.Cells, and saves it as an HTML file in an output folder using HtmlSaveOptions with ExportDocumentProperties disabled, ensuring no document metadata is included.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing XLSX files
            string inputFolder = @"C:\InputXlsx";
            // Output folder for generated HTML files
            string outputFolder = @"C:\OutputHtml";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            try
            {
                // Get all .xlsx files in the input folder (non-recursive)
                string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (string xlsxPath in xlsxFiles)
                {
                    try
                    {
                        // Verify the file exists before loading
                        if (!File.Exists(xlsxPath))
                        {
                            Console.WriteLine($"File not found, skipping: {xlsxPath}");
                            continue;
                        }

                        // Load the workbook from the XLSX file
                        Workbook workbook = new Workbook(xlsxPath);

                        // Configure HTML save options to omit document properties for privacy
                        HtmlSaveOptions saveOptions = new HtmlSaveOptions
                        {
                            ExportDocumentProperties = false, // Do not include document properties in the HTML
                            ExportActiveWorksheetOnly = false // Export all worksheets (adjust as needed)
                        };

                        // Build the output HTML file path (same name, .html extension)
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                        string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                        // Save the workbook as HTML using the configured options
                        workbook.Save(htmlPath, saveOptions);

                        Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(htmlPath)}");
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{xlsxPath}': {exFile.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
