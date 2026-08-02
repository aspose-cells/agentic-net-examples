using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ExportChartWithoutDataSource
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a chart but deliberately do NOT set any data source
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        try
        {
            // Attempt to export the chart to an image file
            // This should throw because the chart has no data source defined
            chart.ToImage("chart.png", ImageType.Png);
            Console.WriteLine("Chart exported successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected path: capture and display the exception message
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }

        // Save the workbook (optional, just to complete the lifecycle)
        workbook.Save("ExportChartWithoutDataSource.xlsx");
    }
}