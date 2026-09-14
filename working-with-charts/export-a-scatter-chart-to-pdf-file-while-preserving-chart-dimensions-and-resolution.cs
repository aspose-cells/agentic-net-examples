// Title: Export a Scatter Chart from an Aspose.Cells Workbook to a PDF while Maintaining Exact Size and Resolution (C#)
// AI Prompts: Generate C# code that creates a scatter chart from worksheet data and saves the workbook as a PDF, preserving the chart's on‑sheet dimensions using Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to keep the chart size unchanged and produce a high‑resolution PDF output. | Demonstrate adding multiple series to a scatter chart and exporting each series accurately to PDF with Aspose.Cells.
// Common Searches: aspnet export scatter chart to pdf preserving chart size Aspose.Cells | C# Aspose.Cells keep chart dimensions when saving workbook as PDF | how to maintain resolution of Excel scatter plot in PDF using Aspose.Cells | PdfSaveOptions OnePagePerSheet false effect on chart size Aspose.Cells
// Tags: scatter chart PDF export Aspose.Cells | preserve chart dimensions PDFSaveOptions | high resolution chart PDF Aspose.Cells | add multiple series scatter chart C# Aspose.Cells | configure onepagepersheet false Aspose.Cells PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills columns A and B with X/Y data, adds a positioned scatter chart, sets a title, and saves the workbook as a PDF using PdfSaveOptions so the chart retains its original size and resolution.
class ExportScatterChartToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the scatter chart (X values in column A, Y values in column B)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 11; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 2);               // X = 0..9
                sheet.Cells[$"B{i}"].PutValue(Math.Sin(i - 2));   // Y = sin(X)
            }

            // Add a scatter chart to the worksheet (position defines size)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Scatter Chart Example";

            // Add a series: X values from A2:A11, Y values from B2:B11
            int seriesIndex = chart.NSeries.Add("B2:B11", true);
            chart.NSeries[seriesIndex].XValues = "A2:A11";

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false  // Keep each sheet on a separate page
            };

            // Save the workbook (including the chart) as a PDF file
            workbook.Save("ScatterChart.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
