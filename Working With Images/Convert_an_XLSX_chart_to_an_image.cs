using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ConvertChartToImage
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Retrieve the first chart in the worksheet
        Chart chart = worksheet.Charts[0];

        // Convert the chart to an image file (PNG format in this example)
        chart.ToImage("chart.png", ImageType.Png);

        Console.WriteLine("Chart has been successfully saved as 'chart.png'.");
    }
}