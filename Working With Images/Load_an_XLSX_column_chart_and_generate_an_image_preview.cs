using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook that contains a column chart
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Retrieve the first chart (assumed to be the column chart)
            Chart chart = worksheet.Charts[0];

            // Export the chart to an image file.
            // The file extension determines the format; here we explicitly set PNG.
            chart.ToImage("chart_preview.png", ImageType.Png);

            Console.WriteLine("Chart preview image saved as 'chart_preview.png'.");
        }
        else
        {
            Console.WriteLine("No charts found in the worksheet.");
        }
    }
}