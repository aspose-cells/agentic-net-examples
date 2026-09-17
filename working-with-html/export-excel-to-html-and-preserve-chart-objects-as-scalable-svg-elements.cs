// Title: Convert an .xlsx workbook to a single HTML page with Base64‑encoded images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file, configures HtmlSaveOptions to embed all images as Base64, and saves the workbook as a single HTML document with Aspose.Cells. | Enhance the code to verify that the source .xlsx file exists and to create the target output folder automatically before invoking the HTML export.
// Common Searches: Aspose.Cells C# export Excel to HTML with embedded Base64 images | how to save an Excel workbook as one HTML file using Aspose.Cells .NET | C# code example for HtmlSaveOptions.ExportImagesAsBase64 in Aspose.Cells | check input file existence before converting Excel to HTML with Aspose.Cells | create output directory automatically when saving HTML from Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export workbook to single HTML Aspose.Cells | embed images as Base64 in Aspose.Cells HTML output | validate input Excel file existence Aspose.Cells | ensure output folder exists before HTML save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an .xlsx workbook, sets HtmlSaveOptions.ExportImagesAsBase64 to embed all images directly in the HTML, ensures the output directory exists, and saves the workbook as a single HTML file while handling missing input files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set HTML save options to embed images as Base64
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true
                // Note: ExportChartImageFormat is not available in this version of Aspose.Cells.
                // Charts will be exported using the default image format.
            };

            // Ensure the output directory exists (if a directory is specified)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
