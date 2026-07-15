using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace PyramidChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: categories (levels) and corresponding values
            sheet.Cells["A1"].PutValue("Level");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Level 1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Level 2");
            sheet.Cells["B3"].PutValue(50);
            sheet.Cells["A4"].PutValue("Level 3");
            sheet.Cells["B4"].PutValue(70);
            sheet.Cells["A5"].PutValue("Level 4");
            sheet.Cells["B5"].PutValue(90);

            // Add a Pyramid chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pyramid, 6, 0, 22, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Ensure each pyramid level gets a distinct color
            // Works because the chart has a single series
            chart.NSeries.IsColorVaried = true;

            // Optional: set a title for clarity
            chart.Title.Text = "Pyramid Chart with Distinct Colors";

            // Save the workbook
            workbook.Save("PyramidChartDistinctColors.xlsx");
        }
    }
}