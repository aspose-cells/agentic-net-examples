// Title: C# – Retrieve and Set a Scatter Chart X‑Axis to a Continuous Numeric Value Axis with Aspose.Cells
// Description: Demonstrates how to create a workbook, add numeric data, insert a scatter chart, obtain the X‑axis via the ValueAxis property, and configure it for linear numeric scaling by disabling logarithmic mode, setting MinValue, MaxValue, MajorUnit, and applying CategoryType.AutomaticScale before saving the file as XLSX.
// Keywords: Aspose.Cells | C# | .NET | chart X axis | value axis | continuous numeric scaling | scatter chart | ValueAxis | CategoryType.AutomaticScale | MinValue | MaxValue | MajorUnit | IsLogarithmic | sample code | GitHub example
// Common Searches: Aspose.Cells set numeric X axis scatter chart C# | retrieve chart ValueAxis Aspose.Cells .NET | change chart X axis to value axis Aspose.Cells | continuous numeric scaling for chart axis Aspose.Cells | Aspose.Cells example X axis MinValue MaxValue
// Developer Intent: Configure the X‑axis of a scatter chart as a linear numeric value axis in Aspose.Cells for .NET.
// Use Cases: Generate a scatter plot from worksheet data where the X‑axis represents sequential numbers with fixed intervals. | Adjust an existing chart’s X‑axis to a linear numeric scale for precise control over range and tick spacing. | Create reusable chart templates that enforce numeric scaling on the X‑axis for time‑series or measurement data.
// AI Prompts: Write C# code using Aspose.Cells to retrieve a chart's X axis and set MinValue, MaxValue, MajorUnit, and IsLogarithmic for continuous numeric scaling. | Explain the steps to change a scatter chart's X axis from category to value axis in Aspose.Cells, including the use of CategoryType.AutomaticScale. | Provide a GitHub‑ready example that creates a workbook, adds numeric data, inserts a scatter chart, and configures the X axis as a linear numeric axis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisDemo
{
    // Demonstrates how to create a workbook, add numeric data, insert a scatter chart, obtain the X‑axis via the ValueAxis property, and configure it for linear numeric scaling by disabling logarithmic mode, setting MinValue, MaxValue, MajorUnit, and applying CategoryType.AutomaticScale before saving the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (numeric categories on X axis)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values: 1,2,3,...
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y values
            }

            // Add a scatter chart (both axes are value axes)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set data source for the series
            chart.NSeries.Add("B2:B10", false);   // Y values
            chart.NSeries[0].XValues = "A2:A10"; // X values

            // Retrieve the X axis (for a scatter chart this is the ValueAxis)
            Axis xAxis = chart.ValueAxis; // X axis in scatter chart

            // Change its type to a continuous numeric scale.
            // For numeric scaling we ensure the axis is treated as a value axis.
            // No explicit enum for "Value", the axis itself is already a value axis,
            // but we can set properties to enforce numeric behavior.
            xAxis.IsLogarithmic = false;          // Linear scale
            xAxis.MinValue = 0;                   // Start at 0
            xAxis.MaxValue = 10;                  // End at 10 (adjust as needed)
            xAxis.MajorUnit = 1;                  // Tick every 1 unit
            xAxis.CategoryType = CategoryType.AutomaticScale; // Ensure automatic scaling

            // Save the workbook
            workbook.Save("ScatterChartWithNumericXAxis.xlsx", SaveFormat.Xlsx);
        }
    }
}
