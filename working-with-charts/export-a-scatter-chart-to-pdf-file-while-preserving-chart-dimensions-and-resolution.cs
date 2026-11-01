// Title: Export a Scatter Chart from Aspose.Cells to PDF while Keeping Original Size and Resolution in C#
// Common Searches: how to save an Aspose.Cells scatter chart as a PDF with original dimensions in C# | Aspose.Cells C# export chart to PDF preserving size and resolution | set PdfSaveOptions to retain chart size when converting Excel to PDF using Aspose.Cells | C# code to create a scatter chart and export it to PDF with Aspose.Cells | preserve chart layout when converting worksheet to PDF with Aspose.Cells
// Tags: scatter chart PDF export Aspose.Cells | chart size preservation PdfSaveOptions | Aspose.Cells convert worksheet to PDF C# | PdfSaveOptions AllColumnsInOnePagePerSheet example | export Excel chart to PDF exact dimensions

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills columns A and B with X/Y data, adds a scatter chart, sets its title, optionally defines its position and size, configures PdfSaveOptions (AllColumnsInOnePagePerSheet) to keep the chart layout, and saves the workbook as a PDF, preserving the chart’s dimensions and resolution.
class ScatterChartToPdf
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
            double[] xValues = { 1, 2, 3, 4, 5 };
            double[] yValues = { 2, 4, 1, 3, 5 };
            for (int i = 0; i < xValues.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(xValues[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(yValues[i]); // Column B
            }

            // Add a scatter chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Sample Scatter Chart";

            // Add a series to the chart using the data range
            int seriesIndex = chart.NSeries.Add("B2:B6", true);
            chart.NSeries[seriesIndex].XValues = "A2:A6";

            // Optional: set chart position and size if needed
            // chart.Position = new Position(5, 0, 0, 0);
            // chart.Height = 300;
            // chart.Width = 500;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Fit the whole sheet on one PDF page
                AllColumnsInOnePagePerSheet = true
            };

            // Define output file path
            string outputPath = "ScatterChart.pdf";

            // Save the workbook (including the chart) as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
