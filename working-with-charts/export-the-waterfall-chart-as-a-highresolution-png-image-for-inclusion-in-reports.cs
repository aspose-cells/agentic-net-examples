// Title: Export Waterfall Chart to High‑Resolution PNG with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a workbook, populate it with waterfall data, add a Waterfall chart, configure ImageOrPrintOptions for 300 DPI PNG output, and save the chart as a high‑quality image while optionally keeping the workbook for later use.
// Keywords: Aspose.Cells | C# | Waterfall chart export | PNG image | high resolution | 300 DPI | ImageOrPrintOptions | chart to image | .NET
// Common Searches: Aspose.Cells export waterfall chart png | C# save chart as high DPI PNG | set chart resolution Aspose.Cells | export specific chart to image .NET | waterfall chart image options Aspose.Cells
// Developer Intent: Generate a 300 DPI PNG file of a Waterfall chart created in an Aspose.Cells workbook.
// Use Cases: Include a crisp waterfall graphic in a financial presentation or report. | Print chart images on high‑quality paper without loss of detail. | Automate batch export of charts for documentation pipelines.
// AI Prompts: Provide C# code to export a Waterfall chart as a 600 DPI PNG using Aspose.Cells. | Show how to loop through all charts on a worksheet and save each as a high‑resolution PNG. | Explain how to adjust image dimensions and DPI when converting a chart to PNG with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace WaterfallChartExport
{
    // Demonstrates how to build a workbook, populate it with waterfall data, add a Waterfall chart, configure ImageOrPrintOptions for 300 DPI PNG output, and save the chart as a high‑quality image while optionally keeping the workbook for later use.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a waterfall chart
            // Column A – Categories, Column B – Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["A5"].PutValue("End");

            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["B2"].PutValue(5000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(-1500);
            sheet.Cells["B5"].PutValue(5500);

            // Add a Waterfall chart (ChartType.Waterfall)
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Optional: set a title
            chart.Title.Text = "Waterfall Chart Example";

            // Configure high‑resolution image options (e.g., 300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,          // Output format
                HorizontalResolution = 300,         // 300 DPI horizontal
                VerticalResolution = 300            // 300 DPI vertical
            };

            // Export the chart to a high‑resolution PNG file
            string imagePath = "WaterfallChart.png";
            chart.ToImage(imagePath, imgOptions);

            // (Optional) Save the workbook for reference
            workbook.Save("WaterfallChartWorkbook.xlsx");

            Console.WriteLine($"Waterfall chart exported to '{imagePath}' with 300 DPI resolution.");
        }
    }
}
