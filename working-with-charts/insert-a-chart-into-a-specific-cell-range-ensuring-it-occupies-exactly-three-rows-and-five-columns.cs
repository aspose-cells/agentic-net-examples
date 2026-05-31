using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart (A1:B5)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Define the cell range where the chart will be placed
            // Top-left corner at row 5, column 2 (zero‑based indices)
            int topRow = 5;          // Row index 5 (6th row in Excel)
            int leftColumn = 2;      // Column index 2 (C column)
            int bottomRow = topRow + 2;   // Occupies 3 rows total
            int rightColumn = leftColumn + 4; // Occupies 5 columns total

            // Add a column chart to the specified range using the Add method
            int chartIndex = sheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("=Sheet1!$A$2:$B$5", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

            // Optionally set a title
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook to a file
            workbook.Save("ChartInRange.xlsx", SaveFormat.Xlsx);
        }
    }
}