// Title: Export a Chart to PNG at 300 DPI Using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds sample data, builds a column chart, sets its data range, configures ImageOrPrintOptions with 300 dpi horizontal and vertical resolution, and saves the chart as a high‑resolution PNG file (Chart_300dpi.png), overwriting any existing file.
// Keywords: Aspose.Cells | C# | chart export | PNG | 300 DPI | ImageOrPrintOptions | high resolution chart | ToImage | Excel chart to image | column chart export
// Common Searches: Aspose.Cells export chart PNG 300 DPI | C# set chart image resolution Aspose.Cells | How to save Excel chart as high‑resolution PNG | ImageOrPrintOptions DPI C# Aspose.Cells | Export column chart to PNG with specific DPI | Aspose.Cells ToImage DPI settings
// Developer Intent: Generate a high‑resolution PNG image of a workbook chart at 300 dpi.
// Use Cases: Print‑ready chart graphics for reports and brochures | High‑quality PNG assets for web dashboards or presentations | Automated batch conversion of multiple Excel charts to 300 dpi PNG files | Embedding chart images in PDF or Word documents with consistent resolution
// AI Prompts: Write C# code to export every chart in an Excel workbook to separate 300 dpi PNG files using Aspose.Cells. | Show how to change the output format to JPEG while keeping a 300 dpi resolution for a chart. | Explain how to programmatically verify the DPI of a PNG file created with ImageOrPrintOptions. | Provide a step‑by‑step guide to export a pie chart to a 300 dpi PNG using Aspose.Cells. | Suggest how to batch process multiple workbooks and export their charts at 300 dpi.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// C# example that creates a workbook, adds sample data, builds a column chart, sets its data range, configures ImageOrPrintOptions with 300 dpi horizontal and vertical resolution, and saves the chart as a high‑resolution PNG file (Chart_300dpi.png), overwriting any existing file.
class ExportChartWithDpi
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Configure image options: PNG format with 300 DPI resolution
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Default format is PNG; explicit setting omitted to avoid missing API issue
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            string outputPath = "Chart_300dpi.png";

            // Ensure we can write the file (overwrite if it exists)
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Export the chart to a PNG file using the specified DPI settings
            chart.ToImage(outputPath, options);

            Console.WriteLine($"Chart exported to '{outputPath}' with 300 DPI resolution.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
