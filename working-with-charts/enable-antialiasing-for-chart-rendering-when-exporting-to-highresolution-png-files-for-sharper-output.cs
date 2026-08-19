// Title: Aspose.Cells .NET – Export Chart to High‑Resolution PNG with Anti‑Aliasing
// Description: Demonstrates how to create a workbook, add a column chart, and export it to a 300 dpi PNG using ImageOrPrintOptions. The high‑resolution setting improves anti‑aliasing for sharper chart rendering, noting that a direct SmoothingMode property is not available in the current API.
// Keywords: Aspose.Cells chart export PNG | high DPI chart image .NET | anti‑aliasing Aspose.Cells | ImageOrPrintOptions resolution | C# chart rendering quality
// Common Searches: Aspose.Cells enable anti‑aliasing for chart PNG | export chart to high resolution PNG C# | set DPI for chart image Aspose.Cells | improve chart image quality Aspose.Cells .NET | smooth chart rendering Aspose.Cells
// Developer Intent: Generate a smoother, high‑resolution PNG of a chart by configuring rendering options in Aspose.Cells for .NET.
// Use Cases: Create a column chart from worksheet data and export it as a 300 dpi PNG for print‑ready reports. | Improve visual clarity of chart lines and text without a dedicated smoothing property. | Save the workbook alongside the exported image while handling possible I/O exceptions.
// AI Prompts: Show C# code to export an Aspose.Cells chart to a high‑resolution PNG with anti‑aliasing. | Explain how DPI settings affect chart anti‑aliasing in Aspose.Cells and what alternatives exist when SmoothingMode is unavailable. | Provide best practices for achieving sharp chart images in Aspose.Cells .NET applications.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, add a column chart, and export it to a 300 dpi PNG using ImageOrPrintOptions. The high‑resolution setting improves anti‑aliasing for sharper chart rendering, noting that a direct SmoothingMode property is not available in the current API.
class ChartAntiAliasingDemo
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

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure image rendering options (default format is PNG)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                HorizontalResolution = 300, // High‑resolution DPI
                VerticalResolution = 300
                // SmoothingMode property is not available in current API version
            };

            // Export the chart to a high‑resolution PNG
            string chartPath = "HighResChart.png";
            try
            {
                chart.ToImage(chartPath, options);
                Console.WriteLine($"Chart image saved to: {Path.GetFullPath(chartPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export chart image: {ex.Message}");
            }

            // Optionally save the workbook
            string workbookPath = "ChartWorkbook.xlsx";
            try
            {
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
