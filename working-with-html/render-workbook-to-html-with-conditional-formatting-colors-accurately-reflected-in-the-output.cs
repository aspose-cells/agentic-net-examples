// Title: Save an Excel workbook as a single HTML page in C# with Aspose.Cells, embedding images as Base64
// Common Searches: C# how to export an .xlsx file to HTML using Aspose.Cells | Aspose.Cells example for saving workbook to HTML with embedded images | convert Excel to a single HTML file in .NET Core with base64 images | Aspose.Cells HtmlSaveOptions usage for HTML export in C# | why conditional formatting colors are not preserved when exporting to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export Excel to HTML Aspose.Cells | single HTML file output from workbook | embed workbook images as Base64 in HTML | conditional formatting export limitation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The sample loads an existing .xlsx workbook with Aspose.Cells, configures HtmlSaveOptions to embed all images as Base64 so the result is a single HTML file, notes that the ExportConditionalFormatting option is unavailable in this version, ensures the output directory exists, saves the HTML, and catches any runtime exceptions.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export images as Base64 to keep a single HTML file
                    ExportImagesAsBase64 = true
                    // Note: ExportConditionalFormatting property is not available in this version of Aspose.Cells
                };

                // Path for the generated HTML file
                string htmlOutputPath = @"C:\Path\To\Output.html";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(htmlOutputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlOutputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully saved as HTML to: {htmlOutputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors (e.g., loading, saving)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
