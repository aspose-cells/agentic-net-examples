// Title: Export a Chart to PDF with Transparent Background using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds sample data, builds a column chart, sets the chart area to fully transparent, and exports the chart directly to a PDF file. Ideal for overlaying the chart on other documents or graphics.
// Keywords: Aspose.Cells | C# | .NET | chart export PDF | transparent chart background | chart area transparency | column chart PDF | Aspose.Cells example | GitHub Aspose.Cells chart | PDF overlay chart
// Common Searches: How to export an Excel chart as a PDF with a transparent background using Aspose.Cells | Aspose.Cells .NET make chart background transparent before PDF export | Export chart to PDF with full transparency Aspose.Cells C# | Aspose.Cells transparent chart PDF example GitHub | C# code to set chart area transparency in Aspose.Cells
// Developer Intent: Create a PDF of a chart that has no background color so it can be layered over other images or PDFs.
// Use Cases: Overlay a chart on a custom slide background for presentations. | Insert a transparent chart into pre‑designed PDF report templates. | Combine charts with web graphics where the chart must blend seamlessly. | Generate reusable chart assets for branding‑consistent marketing materials.
// AI Prompts: Show C# code to set a chart's background to transparent and export it to PDF with Aspose.Cells. | Provide an Aspose.Cells example that exports multiple charts with different transparency levels to separate PDF files. | Explain how to verify that the exported PDF has a fully transparent background using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsChartTransparentPdf
{
    // C# example that creates a workbook, adds sample data, builds a column chart, sets the chart area to fully transparent, and exports the chart directly to a PDF file. Ideal for overlaying the chart on other documents or graphics.
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
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Make the chart background transparent
            chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;
            // Optional: set full transparency (0.0 = opaque, 1.0 = fully transparent)
            chart.ChartArea.Area.Transparency = 1.0;

            // Export the chart to a PDF file with transparent background
            chart.ToPdf("TransparentChart.pdf");

            Console.WriteLine("Chart exported to PDF with transparent background successfully.");
        }
    }
}
