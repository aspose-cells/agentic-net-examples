using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: Workbook constructor)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[$"A{i + 1}"].PutValue($"Cat {i}");
                sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
            }

            // Add a column chart to the worksheet (rule: ChartCollection.Add(ChartType, int, int, int, int))
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 25, 11);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("=Sheet1!$A$1:$B$11", true);

            // Save the workbook (lifecycle rule: Workbook.Save)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}