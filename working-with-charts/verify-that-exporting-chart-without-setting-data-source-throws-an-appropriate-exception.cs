// Title: How to verify that Aspose.Cells Chart.ToImage throws an exception when exporting a chart without a data source in C#
// AI Prompts: Write a C# snippet using Aspose.Cells that creates a chart, intentionally leaves its data source unset, calls Chart.ToImage to export to PNG, and captures the thrown exception. | Generate a unit test in C# for Aspose.Cells that asserts Chart.ToImage raises an error when the chart has no data source.
// Common Searches: Aspose.Cells Chart.ToImage throws exception if chart has no data source | C# test exporting chart without data source Aspose.Cells | how to catch missing data source error when exporting chart image in Aspose.Cells | validate chart data before calling ToImage in Aspose.Cells .NET
// Tags: Aspose.Cells chart export exception | Chart.ToImage missing data source | C# Aspose.Cells chart validation | unit test chart export failure Aspose.Cells | exception handling chart image export

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// // Demonstrates creating a workbook, adding a column chart without assigning a data source, attempting to export the chart to PNG via Chart.ToImage, catching the expected exception, and finally saving the workbook.
class VerifyChartExportWithoutDataSource
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
            // Attempt to export the chart to an image file.
            // Since no data source is defined, Aspose.Cells should throw an exception.
            chart.ToImage("chart.png", ImageType.Png);
            Console.WriteLine("Chart exported successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected path: capture and display the exception message
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }

        // Save the workbook (optional, just to complete the lifecycle)
        workbook.Save("VerifyChartExport.xlsx");
    }
}
