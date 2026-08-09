// Title: Aspose.Cells C# – Set X‑Axis to Value (Continuous Numeric) for Scatter Charts
// Description: Creates a workbook, fills it with numeric X/Y data, adds a scatter chart, retrieves the chart's CategoryAxis, and changes its CategoryType to a numeric/value axis for continuous scaling before saving the file.
// Keywords: Aspose.Cells | C# | scatter chart | X axis | numeric scaling | CategoryAxis | CategoryType.Value | CategoryType.AutomaticScale | chart axis type | continuous axis | Excel chart API
// Common Searches: Aspose.Cells set X axis to value | change chart axis to numeric scaling C# | CategoryAxis CategoryType Value Aspose.Cells | continuous X axis scatter chart Aspose.Cells | how to make X axis treat values as numbers in Excel using Aspose
// Developer Intent: Modify a chart's X‑axis to use a numeric/value scale.
// Use Cases: Generate a scatter plot where the X‑coordinates are treated as continuous numbers. | Convert a category axis to a value axis for time‑series or measurement data. | Update an existing chart after adding rows so the X‑axis reflects proper numeric intervals.
// AI Prompts: Show C# code to retrieve a chart's CategoryAxis and set CategoryType to Value with Aspose.Cells. | Provide an Aspose.Cells example that changes a scatter chart's X axis to continuous numeric scaling and saves the workbook. | Explain the difference between CategoryType.AutomaticScale and CategoryType.Value in Aspose.Cells chart axes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisDemo
{
    // Creates a workbook, fills it with numeric X/Y data, adds a scatter chart, retrieves the chart's CategoryAxis, and changes its CategoryType to a numeric/value axis for continuous scaling before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (numeric X values and Y values)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);          // X: 1,2,3,...
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y: 10,20,30,...
            }

            // Add a scatter chart (suitable for continuous numeric X axis)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set data source for the chart
            chart.NSeries.Add("B2:B10", false);   // Y values
            chart.NSeries[0].XValues = "A2:A10"; // X values

            // Retrieve the X (category) axis and set its type to AutomaticScale
            // This makes the axis treat X values as continuous numeric values
            chart.CategoryAxis.CategoryType = CategoryType.AutomaticScale;

            // Save the workbook
            workbook.Save("ChartWithNumericXAxis.xlsx", SaveFormat.Xlsx);
        }
    }
}
