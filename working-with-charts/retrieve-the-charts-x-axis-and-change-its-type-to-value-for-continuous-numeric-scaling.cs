using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart (numeric X values)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values: 1,2,3,4,5
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y values
            }

            // Add a scatter chart (suitable for numeric X axis)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B6", false);   // Y values
            chart.NSeries[0].XValues = "A2:A6"; // X values

            // Retrieve the X (category) axis
            Axis xAxis = chart.CategoryAxis;

            // Change its type to a continuous numeric scale
            // In Aspose.Cells this is achieved by setting CategoryType to AutomaticScale
            xAxis.CategoryType = CategoryType.AutomaticScale;

            // Save the workbook
            workbook.Save("ChartWithNumericXAxis.xlsx", SaveFormat.Xlsx);
        }
    }
}