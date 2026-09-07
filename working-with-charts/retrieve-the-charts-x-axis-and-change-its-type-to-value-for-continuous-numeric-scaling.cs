// Title: How to set a scatter chart's X‑axis to continuous numeric scaling (AutomaticScale) using Aspose.Cells in C#
// AI Prompts: Generate C# code that retrieves a chart's CategoryAxis and sets its CategoryType to AutomaticScale with Aspose.Cells. | Show me how to configure a scatter chart in Aspose.Cells so the X‑axis is treated as a value axis for numeric data. | Provide a step‑by‑step example of creating a workbook, adding numeric data, inserting a scatter chart, and changing the X‑axis to value scaling before saving.
// Common Searches: Aspose.Cells C# change scatter chart X axis to value axis | set CategoryAxis CategoryType to AutomaticScale in Aspose.Cells example | continuous numeric X axis for chart using Aspose.Cells .NET | how to make X axis numeric in Aspose.Cells scatter plot
// Tags: Aspose.Cells scatter chart X axis AutomaticScale | CategoryAxis CategoryType value scaling C# | numeric X axis Aspose.Cells chart | create chart with continuous X axis .NET | save workbook as Xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills columns A and B with numeric data, adds a scatter chart, assigns Y series and X values, retrieves the chart's CategoryAxis, sets its CategoryType to AutomaticScale for continuous numeric X scaling, and saves the workbook as ChartWithNumericXAxis.xlsx.
class ChangeXAxisToValueScaling
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
            sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values: 1,2,3,...
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y values
        }

        // Add a scatter chart (suitable for continuous numeric X axis)
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 12, 0, 30, 15);
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
