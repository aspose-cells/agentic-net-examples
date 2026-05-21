using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisScalingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
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
            valueAxis.IsAutomaticMinValue = false; // Disable automatic minimum
            valueAxis.IsAutomaticMaxValue = false; // Disable automatic maximum
            valueAxis.MinValue = 0;                 // Set minimum to 0
            valueAxis.MaxValue = 200;               // Set maximum to 200

            // Optional: set major unit for clearer tick marks
            valueAxis.IsAutomaticMajorUnit = false;
            valueAxis.MajorUnit = 20;

            // Save the workbook to a file
            workbook.Save("ChartWithFixedAxisScaling.xlsx");
        }
    }
}