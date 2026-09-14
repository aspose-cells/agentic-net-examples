// Title: Convert a folder of XLSX files to single‑file HTML with Aspose.Cells for .NET while suppressing document properties
// AI Prompts: Generate a C# console application that scans a directory for *.xlsx files and saves each workbook as an HTML file using Aspose.Cells HtmlSaveOptions with ExportDocumentProperties set to false and images embedded as Base64. | Write .NET code to batch process Excel workbooks to HTML, ensuring no document metadata is included and all images are inlined for a single‑file output. | Create a C# script that converts multiple Excel files to privacy‑compliant HTML, logs conversion outcomes, and gracefully handles missing or corrupt files.
// Common Searches: how to batch convert xlsx files to html with aspose.cells c# without exporting document properties | c# aspnet convert multiple excel workbooks to single html files embed images base64 | aspose.cells htmlsaveoptions exportdocumentproperties false example | privacy compliant excel to html conversion using asp.net and aspose.cells | convert all xlsx in a folder to html using aspose.cells console app
// Tags: Aspose.Cells batch XLSX to HTML conversion | HtmlSaveOptions ExportDocumentProperties false | Base64 image embedding in HTML with Aspose.Cells | C# console app bulk Excel to HTML | privacy‑first Excel HTML export

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsxToHtml
{
    // The sample scans a specified input folder, loads each .xlsx workbook with Aspose.Cells, and saves it as a single HTML file using HtmlSaveOptions that disable document property export and embed images as Base64, while creating the output directory, handling missing files, and logging conversion results.
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
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the input folder (non-recursive)
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            // HtmlSaveOptions with document properties omitted for privacy compliance
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDocumentProperties = false, // Do not include document properties in the HTML output
                ExportImagesAsBase64 = true        // Embed images as base64 to keep a single HTML file
            };

            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Ensure the source file exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Load the workbook from the XLSX file
                    Workbook workbook = new Workbook(xlsxPath);

                    // Determine output HTML file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Converted: {xlsxPath} -> {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
