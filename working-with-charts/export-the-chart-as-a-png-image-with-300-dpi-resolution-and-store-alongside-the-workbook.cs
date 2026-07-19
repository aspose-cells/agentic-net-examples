// Title: Export Aspose.Cells Chart to PNG at 300 DPI and Save Workbook (C#)
// Description: Creates a new workbook, fills sample data, adds a column chart, configures ImageOrPrintOptions for PNG with 300 dpi horizontal and vertical resolution, exports the chart as "ChartImage.png", and saves the workbook as "Workbook.xlsx" in the same directory.
// Keywords: Aspose.Cells | C# chart export | PNG 300 DPI | ImageOrPrintOptions | export chart as image | .NET Excel chart | high‑resolution chart image | save workbook and chart
// Common Searches: Aspose.Cells export chart PNG 300 DPI C# | how to save chart image with workbook using Aspose.Cells | set DPI for chart export Aspose.Cells | C# code to export Excel chart as high‑resolution PNG
// Developer Intent: Generate a 300 dpi PNG file from an Aspose.Cells chart and keep the Excel workbook in the same folder.
// Use Cases: Create print‑ready chart graphics for PDF reports. | Attach high‑quality PNG snapshots of charts to emails while preserving the source workbook. | Produce marketing assets that require 300 dpi chart images alongside the original Excel file.
// AI Prompts: Show C# code to export multiple Aspose.Cells charts to separate PNG files, each with a custom DPI. | Demonstrate how to export an Aspose.Cells chart as a JPEG image at 600 dpi. | Explain how to embed the exported PNG chart into a Word document using Aspose.Words after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartExport
{
    // Creates a new workbook, fills sample data, adds a column chart, configures ImageOrPrintOptions for PNG with 300 dpi horizontal and vertical resolution, exports the chart as "ChartImage.png", and saves the workbook as "Workbook.xlsx" in the same directory.
    class Program
    {
        static void Main()
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

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Configure image options: PNG format with 300 DPI resolution
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = ImageType.Png;
            options.HorizontalResolution = 300;
            options.VerticalResolution = 300;

            // Export the chart as a PNG image with the specified DPI
            chart.ToImage("ChartImage.png", options);

            // Save the workbook in the same folder
            workbook.Save("Workbook.xlsx");

            Console.WriteLine("Chart exported as PNG with 300 DPI and workbook saved.");
        }
    }
}
