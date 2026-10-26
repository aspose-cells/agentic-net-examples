// Title: Save an Excel workbook with a pivot table as an HTML page using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook containing pivot table to HTML | How to save Excel file with pivot tables as a single HTML file in .NET | How to embed images as Base64 when exporting Excel to HTML with Aspose.Cells | Export all worksheets to HTML preserving pivot table layout using Aspose.Cells | Convert Excel workbook to HTML with pivot tables and base64 images in C#
// Tags: HtmlSaveOptions ExportActiveWorksheetOnly false Aspose.Cells | pivot table HTML export Aspose.Cells | embed images as Base64 Aspose.Cells HTML | combine multiple worksheets into one HTML file Aspose.Cells | preserve pivot table layout during Excel to HTML conversion C#

using System;
using System.IO;
using System.Drawing.Imaging;
using Aspose.Cells;

// The example loads an existing workbook that contains a pivot table, configures HtmlSaveOptions to export all worksheets and embed images as Base64 strings, and then saves the workbook as a single HTML page, handling missing files and runtime exceptions.
class PivotTableHtmlExport
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the pivot table.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets.
                ExportActiveWorksheetOnly = false,

                // Preserve images as Base64 strings.
                ExportImagesAsBase64 = true

                // Note: ExportChartImageFormat is not required; default format is PNG.
            };

            // Save the workbook as an HTML page.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
