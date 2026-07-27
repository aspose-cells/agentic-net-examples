// Title: C# – Set Chart X‑Axis to a Value Axis (AutomaticScale) with Aspose.Cells
// Description: Creates a workbook, fills columns A and B with numeric data, adds a scatter chart, links Y values to B2:B6 and X values to A2:A6, retrieves the chart's CategoryAxis, changes its CategoryType to AutomaticScale for continuous numeric scaling, and saves the result as ChartXAxisValueType.xlsx.
// Keywords: Aspose.Cells | C# chart axis | CategoryAxis | AutomaticScale | value axis | continuous numeric scaling | scatter chart | Excel chart axis type | CategoryType.AutomaticScale | chart scaling Aspose
// Common Searches: Aspose.Cells change X axis to value axis | CategoryAxis AutomaticScale example C# | set chart axis type to numeric Aspose.Cells | continuous X axis scaling in Excel with Aspose | retrieve and modify chart axis properties C#
// Developer Intent: Convert a chart's category X‑axis to a numeric value axis for continuous scaling.
// Use Cases: Generate a scatter chart where the X values are treated as a numeric scale rather than discrete categories. | Create time‑series or measurement charts that require a continuous X‑axis in an Excel file produced programmatically. | Adjust existing Excel charts to improve data visualization by switching the X‑axis to AutomaticScale.
// AI Prompts: Show C# code that retrieves a chart's CategoryAxis and sets CategoryType to AutomaticScale using Aspose.Cells. | How can I change a line or scatter chart X‑axis from category to value axis in Aspose.Cells for .NET? | Explain the impact of CategoryType.AutomaticScale on Excel chart axes generated with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills columns A and B with numeric data, adds a scatter chart, links Y values to B2:B6 and X values to A2:A6, retrieves the chart's CategoryAxis, changes its CategoryType to AutomaticScale for continuous numeric scaling, and saves the result as ChartXAxisValueType.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (numeric X values for continuous scaling)
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values: 1,2,3,4,5
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y values
        }

        // Add a scatter chart (X axis is a value axis by nature)
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set Y values and link X values to the numeric range
        chart.NSeries.Add("B2:B6", false);
        chart.NSeries[0].XValues = "A2:A6";

        // Retrieve the X (category) axis of the chart
        Axis xAxis = chart.CategoryAxis;

        // Change its type to continuous numeric scaling
        // AutomaticScale lets Excel treat the axis as a numeric value axis
        xAxis.CategoryType = CategoryType.AutomaticScale;

        // Save the workbook
        workbook.Save("ChartXAxisValueType.xlsx", SaveFormat.Xlsx);
    }
}
