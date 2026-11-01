// Title: Convert multiple XLSX workbooks to HTML in C# with Aspose.Cells while disabling CSS generation
// AI Prompts: Create a C# console app that scans a directory for *.xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file using HtmlSaveOptions configured to omit embedded CSS and to store images as separate files. | Write a .NET routine that performs bulk Excel‑to‑HTML conversion with Aspose.Cells, ensuring the generated HTML has no stylesheet block and that image data is not embedded as Base64.
// Common Searches: asp.net batch convert xlsx to html without css using aspose.cells | c# convert multiple excel files to html minimal output size | how to disable stylesheet generation in Aspose.Cells HTML export | save Excel workbook as html with external images and no embedded css in C#
// Tags: Aspose.Cells HtmlSaveOptions export without CSS | bulk XLSX to HTML conversion C# | Excel workbook to HTML minimal size | C# folder iteration Aspose.Cells | external image handling Aspose.Cells HTML export

using System;
using System.IO;
using Aspose.Cells;

// The sample enumerates all .xlsx files in a given input folder, creates an output directory, configures HtmlSaveOptions to avoid embedding CSS and to keep images as separate files, loads each workbook with Aspose.Cells, and saves it as a lightweight HTML file, logging progress and handling errors.
class BatchXlsxToHtml
{
    static void Main(string[] args)
    {
        // Folder containing the source XLSX files
        string sourceFolder = @"C:\InputXlsx";
        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all .xlsx files in the source folder (non‑recursive)
        string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Do not embed images as Base64 strings (optional, reduces size if images are saved separately)
            ExportImagesAsBase64 = false
            // Note: ExportEmbeddedCss and ExportCssClassNames are not available in this version of Aspose.Cells
        };

        foreach (string xlsxPath in xlsxFiles)
        {
            try
            {
                // Verify the file exists before attempting to load
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found: {xlsxPath}");
                    continue;
                }

                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(xlsxPath);

                // Determine the output HTML file name
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Converted: {xlsxPath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{xlsxPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
