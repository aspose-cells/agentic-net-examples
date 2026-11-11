// Title: Export Aspose.Cells workbook to HTML with PNG chart images and high‑resolution rendering in C#
// Common Searches: C# Aspose.Cells save workbook as HTML with chart images in PNG format | How to set HtmlSaveOptions ExportChartImageFormat to PNG for high‑resolution charts | Aspose.Cells generate HTML output with high quality chart images | Saving Excel charts as PNG when exporting to HTML using Aspose.Cells .NET | Increase chart image resolution in HTML export with Aspose.Cells
// Tags: HtmlSaveOptions chart image format PNG | Aspose.Cells HTML chart rendering | C# export workbook to HTML with charts | high‑resolution chart output Aspose.Cells | save Excel charts as PNG in HTML

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data and a column chart, configures HtmlSaveOptions.ExportChartImageFormat to PNG to produce high‑resolution chart images, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(15);
            sheet.Cells["B2"].PutValue(25);
            sheet.Cells["B3"].PutValue(35);

            // Insert a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add first series and set its name
            int seriesIdx1 = chart.NSeries.Add("A1:A3", true);
            chart.NSeries[seriesIdx1].Name = "Series1";

            // Add second series and set its name
            int seriesIdx2 = chart.NSeries.Add("B1:B3", true);
            chart.NSeries[seriesIdx2].Name = "Series2";

            // Configure HTML save options (default settings are sufficient for chart images)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
