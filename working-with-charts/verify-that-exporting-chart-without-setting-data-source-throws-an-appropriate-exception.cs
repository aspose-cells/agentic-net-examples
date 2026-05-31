using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExportTest
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a chart to the worksheet but do NOT set any data source
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Attempt to export the chart to an image without a data source
            try
            {
                // This should throw an exception because the chart has no data range
                chart.ToImage("ChartWithoutData.png", ImageType.Png);
                Console.WriteLine("Unexpected: Chart exported without data source.");
            }
            catch (Exception ex)
            {
                // Expected path: capture and display the exception details
                Console.WriteLine("Expected exception caught while exporting chart without data source:");
                Console.WriteLine(ex.Message);
            }

            // Attempt to export the chart to a PDF without a data source
            try
            {
                chart.ToPdf("ChartWithoutData.pdf");
                Console.WriteLine("Unexpected: Chart exported to PDF without data source.");
            }
            catch (Exception ex)
            {
                // Expected path: capture and display the exception details
                Console.WriteLine("Expected exception caught while exporting chart to PDF without data source:");
                Console.WriteLine(ex.Message);
            }

            // Save the workbook (lifecycle save) – the workbook itself is valid even without chart data
            workbook.Save("WorkbookWithEmptyChart.xlsx");
        }
    }
}