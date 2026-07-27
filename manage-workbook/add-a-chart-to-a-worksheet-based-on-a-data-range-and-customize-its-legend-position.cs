using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            // Column A: Categories, Column B: Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);

            // Add a column chart to the worksheet.
            // Parameters: chart type, data range, plot by column (true), top row, left column, bottom row, right column
            int chartIndex = sheet.Charts.Add(
                ChartType.Column,          // Chart type
                "A1:B5",                   // Data range (including headers)
                true,                      // Plot by column
                7,                         // Top row position of the chart
                1,                         // Left column position of the chart
                25,                        // Bottom row position of the chart
                10);                       // Right column position of the chart

            // Retrieve the newly created chart
            Chart chart = sheet.Charts[chartIndex];

            // Optionally, set the chart title
            chart.Title.Text = "Sample Column Chart";

            // Customize the legend position (e.g., place it at the bottom of the chart)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook to an XLSX file
            workbook.Save("ChartWithCustomLegend.xlsx", SaveFormat.Xlsx);
        }
    }
}