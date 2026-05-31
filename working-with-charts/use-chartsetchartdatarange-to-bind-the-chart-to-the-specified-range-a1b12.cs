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

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Populate sample data in the range A1:B12
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 12; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the specified data range A1:B12 (vertical series)
            chart.SetChartDataRange("A1:B12", true);

            // Save the workbook
            workbook.Save("ChartWithDataRange.xlsx");
        }
    }
}