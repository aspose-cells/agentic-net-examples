// Title: Aspose.Cells C# – Export Chart Without Data Source Throws Exception
// Description: Shows how calling Chart.ToImage on a column chart that has no data source in Aspose.Cells triggers an exception. The sample creates a workbook, adds an empty chart, attempts PNG export, catches the expected error, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | Chart.ToImage | missing data source | exception handling | export chart | column chart | chart validation | chart export error
// Common Searches: Aspose.Cells Chart.ToImage exception | export chart without data source Aspose.Cells | how to catch missing data source error in Aspose.Cells | what exception is thrown when chart has no data source | validate chart data before exporting Aspose.Cells
// Developer Intent: Confirm that exporting a chart with no data source using Aspose.Cells raises the appropriate exception.
// Use Cases: Write a unit test that asserts an exception is thrown when Chart.ToImage is called on an empty chart. | Add pre‑export validation to reporting code to avoid runtime failures caused by missing chart data. | Log and present user‑friendly messages when a chart cannot be exported because its data source is undefined.
// AI Prompts: Generate an xUnit test that verifies Aspose.Cells throws the correct exception when exporting a chart without a data source. | Provide C# code that checks if a chart has a data source and throws a custom InvalidOperationException before calling ToImage. | Explain which specific Aspose.Cells exception type is raised for a missing chart data source and show best‑practice handling patterns.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExportTest
{
    // Shows how calling Chart.ToImage on a column chart that has no data source in Aspose.Cells triggers an exception. The sample creates a workbook, adds an empty chart, attempts PNG export, catches the expected error, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a chart without setting any data source
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            try
            {
                // Attempt to export the chart to an image file.
                // Since no data source is defined, Aspose.Cells should throw an exception.
                chart.ToImage("ChartWithoutData.png", ImageType.Png);
                Console.WriteLine("Chart exported successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected path: an exception indicating missing data source.
                Console.WriteLine("Expected exception caught: " + ex.Message);
            }

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("ChartExportTest.xlsx");
        }
    }
}
