// Title: Batch convert multiple XLSX workbooks to compact HTML with unused style removal using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, calls RemoveUnusedStyles, and saves the result as a single‑file HTML with images embedded as Base64. | Write C# code that uses Aspose.Cells save options to export only the active worksheet and produce minimal HTML while processing a directory of Excel files. | Create a robust batch conversion utility in C# that validates input and output directories, handles missing files, and converts each Excel workbook to reduced‑size HTML using Aspose.Cells.
// Common Searches: aspnet batch convert xlsx to html with aspose.cells removeunusedstyles | how to export excel worksheets as single html file with base64 images using c# | c# iterate through directory of excel files and save each as minimal html using aspose | reduce html size when converting excel to html with aspose.cells saveoptions | remove unused styles from workbook before html export asp.net
// Tags: Aspose.Cells unused style cleanup | C# process Excel files in a folder | HtmlSaveOptions active worksheet export | Base64 image embedding Aspose.Cells | compact HTML generation from Excel | multiple workbook HTML export

using System;
using System.IO;
using Aspose.Cells;

// // This C# console application scans a specified input directory for .xlsx files, loads each workbook with Aspose.Cells, removes unused styles to shrink size, configures HTML save options to export only the active worksheet and embed images as Base64, then saves a compact single‑file HTML output to a target directory while handling errors and validating paths.
class BatchXlsxToHtml
{
    static void Main()
    {
        // Directory containing the source XLSX files
        string inputDirectory = @"C:\InputXlsx";

        // Directory where the reduced‑size HTML files will be saved
        string outputDirectory = @"C:\OutputHtml";

        // Verify input directory exists
        if (!Directory.Exists(inputDirectory))
        {
            Console.WriteLine($"Input directory does not exist: {inputDirectory}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Get all XLSX files in the input directory
        string[] xlsxFiles = Directory.GetFiles(inputDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

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

                // Remove unused styles to reduce the file size (compatible with all versions)
                workbook.RemoveUnusedStyles();

                // Configure HTML save options for minimal output
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export only the active worksheet (optional, further reduces size)
                    ExportActiveWorksheetOnly = true,

                    // Export images as Base64 strings to keep a single HTML file
                    ExportImagesAsBase64 = true
                };

                // Determine the output HTML file name
                string htmlFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".html";
                string htmlPath = Path.Combine(outputDirectory, htmlFileName);

                // Save the workbook as a reduced‑size HTML file
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {htmlFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed. HTML files are saved to: " + outputDirectory);
    }
}
