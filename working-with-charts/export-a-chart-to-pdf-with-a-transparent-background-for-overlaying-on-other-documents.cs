// Title: Export Aspose.Cells chart to PDF with a transparent background using C#
// Description: Creates a workbook, adds sample data, builds a column chart, sets the chart area to fully transparent, and saves the chart directly as a PDF file (TransparentChart.pdf) for seamless overlay on other documents.
// Keywords: Aspose.Cells chart export PDF | transparent chart background C# | ChartArea.BackgroundMode Transparent | Aspose.Cells ToPdf transparency | C# export chart as PDF | PDF overlay chart Aspose
// Common Searches: Aspose.Cells export chart to PDF transparent background | C# make chart background transparent before PDF export | How to save Aspose.Cells chart as PDF with no background | Transparent PDF chart Aspose.Cells example
// Developer Intent: Generate a PDF version of a chart with a completely transparent background for compositing with other PDFs or graphics.
// Use Cases: Overlay a chart on a pre‑designed PDF template without a white box. | Create transparent charts for web dashboards that can be layered over images. | Combine multiple PDF visualizations by placing transparent charts on existing pages.
// AI Prompts: Show C# code that creates an Aspose.Cells chart, sets its background to fully transparent, and exports it to PDF. | Provide an example of exporting a transparent Aspose.Cells chart to PDF for overlay purposes. | Explain how to adjust chart area and legend transparency when saving a chart as PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTransparentPdf
{
    // Creates a workbook, adds sample data, builds a column chart, sets the chart area to fully transparent, and saves the chart directly as a PDF file (TransparentChart.pdf) for seamless overlay on other documents.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Make the chart background transparent
            chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;
            // Optional: ensure the chart area itself is fully transparent
            chart.ChartArea.Area.Transparency = 1.0; // 1.0 = completely transparent

            // Export the chart to a PDF file with transparent background
            chart.ToPdf("TransparentChart.pdf");

            Console.WriteLine("Chart exported to TransparentChart.pdf with transparent background.");
        }
    }
}
