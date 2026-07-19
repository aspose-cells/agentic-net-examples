// Title: C# Aspose.Cells – Bind a Chart to an OFFSET Dynamic Named Range
// Description: Demonstrates how to create a workbook, fill columns with categories and values, define a named range using an OFFSET formula that grows with new rows, add a column chart, and bind the chart series to that dynamic range so the chart updates automatically.
// Keywords: Aspose.Cells OFFSET named range | dynamic chart range C# | bind chart to named range .NET | auto‑expanding chart data Aspose.Cells | Excel chart dynamic source C# | Aspose.Cells chart binding
// Common Searches: Aspose.Cells create OFFSET named range .NET | bind chart to dynamic range C# Aspose.Cells | auto expand chart data when adding rows Aspose.Cells | set chart categories from named range C# | dynamic Excel chart with Aspose.Cells
// Developer Intent: Generate a chart that automatically reflects added rows by linking it to an OFFSET‑based named range.
// Use Cases: Sales report where the chart grows with new sales entries. | Dashboard displaying monthly metrics that expands as new months are added. | Template workbook that updates its chart as users fill additional data rows.
// AI Prompts: Show how to modify the OFFSET formula to make the column count dynamic in Aspose.Cells. | Provide C# code to refresh a chart after programmatically inserting new rows. | Explain how to reuse the same dynamic named range across multiple charts in one workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, fill columns with categories and values, define a named range using an OFFSET formula that grows with new rows, add a column chart, and bind the chart series to that dynamic range so the chart updates automatically.
class BindChartToOffsetNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in columns A (Category) and B (Value)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 10; i++)
        {
            sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
            sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
        }

        // Create an OFFSET‑based named range that expands automatically with new rows
        // The range starts at A1 and includes two columns (A and B) and as many rows as there are entries in column A
        int nameIdx = workbook.Worksheets.Names.Add("ChartData");
        Name chartDataName = workbook.Worksheets.Names[nameIdx];
        chartDataName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),2)";

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIdx];

        // Bind the chart to the OFFSET‑based named range
        // The first column of the range will be used for categories, the second for values
        chart.NSeries.Add("=ChartData", true);

        // Optional: set chart title
        chart.Title.Text = "Dynamic Data Chart";

        // Save the workbook
        workbook.Save("DynamicChart.xlsx");
    }
}
