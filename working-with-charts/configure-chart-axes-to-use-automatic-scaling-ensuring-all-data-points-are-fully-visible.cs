using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoScalingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the value axis uses automatic scaling for min and max values
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsAutomaticMinValue = true;   // Let Excel determine the minimum automatically
            valueAxis.IsAutomaticMaxValue = true;   // Let Excel determine the maximum automatically
            valueAxis.IsAutomaticMajorUnit = true;  // Automatic major unit
            valueAxis.IsAutomaticMinorUnit = true;  // Automatic minor unit

            // Ensure the category axis also uses automatic scaling (if applicable)
            Axis categoryAxis = chart.CategoryAxis;
            categoryAxis.IsAutomaticMinValue = true;
            categoryAxis.IsAutomaticMaxValue = true;
            categoryAxis.IsAutomaticMajorUnit = true;
            categoryAxis.IsAutomaticMinorUnit = true;

            // Recalculate the chart to apply automatic positioning and scaling
            chart.Calculate();

            // Save the workbook with the configured chart
            workbook.Save("ChartAutoScalingDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}