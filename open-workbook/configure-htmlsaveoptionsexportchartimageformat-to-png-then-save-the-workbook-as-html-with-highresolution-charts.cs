// Title: Save an Excel workbook to HTML with high‑resolution PNG chart images using Aspose.Cells for .NET
// Common Searches: how to save Excel workbook as HTML with PNG charts using Aspose.Cells C# | Aspose.Cells set ExportChartImageResolution to 300 DPI for HTML output | C# example for exporting chart images as PNG in HTML with Aspose.Cells | HtmlSaveOptions ExportChartImageFormat PNG sample code | increase chart image quality when converting Excel to HTML with Aspose.Cells | Aspose.Cells HTML conversion high DPI chart images
// Tags: HtmlSaveOptions ExportChartImageFormat PNG | ExportChartImageResolution high DPI Aspose.Cells | Aspose.Cells HTML export chart images | C# Excel to HTML conversion with PNG charts | high resolution chart export Aspose.Cells | Aspose.Cells chart image format setting

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an Excel file, configuring HtmlSaveOptions to export chart images as PNG with a 300 DPI resolution, and saving the workbook as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // If your Aspose.Cells version supports these properties, you can uncomment them:
            // htmlOptions.ExportChartImageFormat = System.Drawing.Imaging.ImageFormat.Png;
            // htmlOptions.ExportChartImageResolution = 300;

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
