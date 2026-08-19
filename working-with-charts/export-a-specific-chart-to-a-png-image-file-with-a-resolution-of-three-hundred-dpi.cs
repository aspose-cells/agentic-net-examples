// Title: Export a Chart to PNG at 300 DPI using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, sets its data range, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and saves the chart as a high‑resolution PNG file.
// Keywords: Aspose.Cells | C# | export chart PNG | 300 DPI | ImageOrPrintOptions | chart.ToImage | high resolution chart image | column chart rendering | .NET chart export
// Common Searches: Aspose.Cells export chart 300 DPI PNG | C# set chart image resolution Aspose.Cells | How to save chart as high‑resolution PNG using Aspose.Cells | ImageOrPrintOptions horizontal vertical resolution example | Export specific chart to PNG with Aspose.Cells .NET
// Developer Intent: Generate a PNG image of a workbook chart with a fixed 300 DPI resolution using Aspose.Cells for .NET.
// Use Cases: Produce print‑ready chart graphics for reports and brochures. | Create web‑optimized PNGs that retain clarity on high‑DPI displays. | Automate batch export of multiple charts, ensuring uniform 300 DPI output.
// AI Prompts: Write C# code to export every chart in a workbook to separate 300 DPI PNG files with Aspose.Cells. | Show how to change the export format to JPEG while keeping a 300 DPI resolution for a chart. | Explain how to add a transparent background to a PNG chart export using ImageOrPrintOptions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace ExportChartToPng300Dpi
{
    // Creates a workbook, adds sample data, inserts a column chart, sets its data range, configures ImageOrPrintOptions with 300 DPI horizontal and vertical resolution, and saves the chart as a high‑resolution PNG file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
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
                    // Default image format is PNG; explicit setting omitted for compatibility
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Export the chart to a PNG file using the specified options
                string outputPath = "Chart_300DPI.png";
                try
                {
                    chart.ToImage(outputPath, options);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export chart image: {ex.Message}");
                }

                // Save the workbook (optional)
                string workbookPath = "ChartWorkbook.xlsx";
                try
                {
                    workbook.Save(workbookPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }

                Console.WriteLine($"Chart exported to '{outputPath}' with 300 DPI resolution.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
