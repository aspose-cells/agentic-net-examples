// Title: Aspose.Cells .NET – Verify exception when exporting a chart without a data source
// Description: C# sample that creates a workbook, adds a column chart with no series, and calls Chart.ToImage. The code catches the expected exception, confirming that exporting a chart without a data source fails, and then saves the workbook.
// Keywords: Aspose.Cells | .NET | Chart.ToImage | chart export exception | no data series | InvalidOperationException | C# chart export error
// Common Searches: Aspose.Cells exception when exporting chart with no data | Chart.ToImage throws error without series | How to test chart export failure in Aspose.Cells | Validate chart data source before calling ToImage | C# Aspose.Cells chart export without data source
// Developer Intent: Confirm that calling Chart.ToImage on a chart that has no data series raises an exception.
// Use Cases: Create a unit test that asserts an exception is thrown for empty‑series charts. | Wrap chart export in try‑catch to log missing data source errors. | Check chart.Series.Count before export and provide a custom error message.
// AI Prompts: Generate an NUnit test that verifies Aspose.Cells throws InvalidOperationException when Chart.ToImage is called on a chart without any series. | Provide C# code that inspects a Chart object for existing series and throws a custom ChartDataMissingException before exporting. | Explain how to capture and log the exception message from Chart.ToImage when the chart lacks a data source.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExportTest
{
    // C# sample that creates a workbook, adds a column chart with no series, and calls Chart.ToImage. The code catches the expected exception, confirming that exporting a chart without a data source fails, and then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a chart to the worksheet but do NOT set any data source
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Attempt to export the chart to an image file.
            // Expect an exception because the chart has no data source.
            try
            {
                chart.ToImage("ChartWithoutData.png", ImageType.Png);
                Console.WriteLine("Export succeeded unexpectedly. No exception was thrown.");
            }
            catch (Exception ex)
            {
                // Verify that an exception is thrown and display its message
                Console.WriteLine("Expected exception caught: " + ex.Message);
            }

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("ChartExportTest.xlsx");
        }
    }
}
