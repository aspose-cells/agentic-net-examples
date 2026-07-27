using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisScalingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(120);
            sheet.Cells["B4"].PutValue(190);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data series for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the vertical (value) axis scaling
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsAutomaticMinValue = false; // disable automatic min
            valueAxis.IsAutomaticMaxValue = false; // disable automatic max
            valueAxis.MinValue = 0;                 // set minimum to 0
            valueAxis.MaxValue = 200;               // set maximum to 200

            // Save the workbook
            workbook.Save("ChartWithFixedAxis.xlsx");
        }
    }
}