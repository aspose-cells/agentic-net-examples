// Title: C# – Export Aspose.Cells Chart to PDF with Transparent Background
// Description: Demonstrates how to create a workbook, add a column chart, set the chart area to transparent, and save only the chart as a PDF using Aspose.Cells' Chart.ToPdf method. Ideal for generating overlay‑ready graphics.
// Keywords: Aspose.Cells C# chart PDF export | transparent chart background | Chart.ToPdf transparent | export chart as PDF Aspose | C# Aspose.Cells transparent PDF | overlay chart PDF
// Common Searches: Aspose.Cells export chart to PDF transparent | C# make chart background transparent Aspose | Chart.ToPdf transparent background example | save Aspose.Cells chart as PDF without background | how to overlay Aspose chart on other PDFs
// Developer Intent: Generate a PDF file that contains only a chart from an Aspose.Cells workbook, preserving a fully transparent background for compositing.
// Use Cases: Create overlayable charts for financial reports or dashboards. | Insert clean, background‑free graphics into presentations or PDFs. | Produce watermark‑free visual elements for branding or marketing materials.
// AI Prompts: Show C# code to export an Aspose.Cells chart to PNG with a transparent background. | Explain how to set a chart's background to transparent and customize the PDF page size in Aspose.Cells. | Guide me on merging multiple transparent‑background chart PDFs into one document using Aspose.Pdf.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTransparentPdf
{
    // Demonstrates how to create a workbook, add a column chart, set the chart area to transparent, and save only the chart as a PDF using Aspose.Cells' Chart.ToPdf method. Ideal for generating overlay‑ready graphics.
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Make the chart background transparent
            chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;

            // Export the chart to a PDF file with transparent background
            chart.ToPdf("TransparentChart.pdf");

            Console.WriteLine("Chart exported to PDF with transparent background successfully.");
        }
    }
}
