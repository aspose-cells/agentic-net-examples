using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing XLSX file that contains a pie chart
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet has at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart (assumed to be the pie chart)
            Chart chart = worksheet.Charts[0];

            // Generate an image preview of the chart and save it as PNG
            // This uses the Chart.ToImage(string, ImageType) method (rule-compliant)
            chart.ToImage("chart_preview.png", ImageType.Png);
        }
        else
        {
            Console.WriteLine("No charts found in the worksheet.");
        }
    }
}