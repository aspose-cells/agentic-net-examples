// Title: Convert an Excel workbook to self-contained HTML with original column widths using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML preserving column widths | How to keep Excel column widths when saving as HTML with Aspose.Cells | Generate HTML from .xlsx with CSS width styles using Aspose.Cells .NET | C# save Excel as HTML with embedded Base64 images and column width CSS | Aspose.Cells HtmlSaveOptions to maintain layout in HTML output
// Tags: Aspose.Cells HtmlSaveOptions column width preservation | C# export Excel to HTML self-contained | HTML output with Base64 images Aspose.Cells | CSS width styling for Excel columns Aspose.Cells | SaveFormat.Html Aspose.Cells configuration

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to export all worksheets, embed images as Base64, and retain the original column widths via generated CSS width properties, then saves the result as a self‑contained HTML file.
class ExcelToHtmlExporter
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = @"C:\Path\To\YourWorkbook.xlsx";

        // Verify that the Excel file exists to avoid FileNotFoundException
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The Excel file was not found at '{excelPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets (set to true if you need only the active sheet)
                ExportActiveWorksheetOnly = false,

                // Embed images as Base64 strings to keep the HTML self‑contained
                ExportImagesAsBase64 = true
            };

            // Path for the output HTML file
            string htmlPath = @"C:\Path\To\ExportedWorkbook.html";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(htmlPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine("Excel file has been successfully exported to HTML with original column widths preserved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during export: {ex.Message}");
        }
    }
}
