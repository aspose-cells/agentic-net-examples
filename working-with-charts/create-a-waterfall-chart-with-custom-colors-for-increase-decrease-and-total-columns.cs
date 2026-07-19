// Title: Aspose.Cells C# – Create a Waterfall Chart with Custom Colors for Increase, Decrease, and Totals
// Description: Demonstrates how to build a Waterfall chart in a new workbook using Aspose.Cells for .NET, bind category and value ranges, and assign distinct foreground colors (gray for totals, green for increases, red for decreases) before saving the file as WaterfallCustomColors.xlsx.
// Keywords: Aspose.Cells waterfall chart | C# custom waterfall colors | set point color Aspose.Cells | waterfall chart increase decrease total | Aspose.Cells chart formatting .NET | waterfall chart point foreground color | Aspose.Cells example C#
// Common Searches: Aspose.Cells change waterfall column colors C# | how to set increase and decrease colors in Aspose.Cells waterfall chart | custom total column color Aspose.Cells .NET | waterfall chart point color formatting Aspose | C# Aspose.Cells waterfall chart example
// Developer Intent: Generate a Waterfall chart and apply specific colors to total, increase, and decrease columns using Aspose.Cells for .NET.
// Use Cases: Financial statement visualizations where start/end totals are neutral (gray) and revenue/cost changes are highlighted in green and red. | Sales pipeline dashboards that emphasize gains and losses with custom colors while keeping overall totals subdued. | Project budget waterfall reports that differentiate allocated funds, overruns, and final balance through distinct column colors.
// AI Prompts: Show how to replace the predefined System.Drawing colors with custom RGB values in the waterfall chart example. | Provide code to add data labels to each waterfall column while preserving the custom point colors. | Explain how to apply the same custom color scheme to multiple series in a single waterfall chart using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to build a Waterfall chart in a new workbook using Aspose.Cells for .NET, bind category and value ranges, and assign distinct foreground colors (gray for totals, green for increases, red for decreases) before saving the file as WaterfallCustomColors.xlsx.
class WaterfallChartWithCustomColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data for the waterfall chart
        // Column A: Categories, Column B: Values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        sheet.Cells["A2"].PutValue("Start");
        sheet.Cells["B2"].PutValue(100);   // Total (start)

        sheet.Cells["A3"].PutValue("Revenue");
        sheet.Cells["B3"].PutValue(150);   // Increase

        sheet.Cells["A4"].PutValue("Cost");
        sheet.Cells["B4"].PutValue(-30);   // Decrease

        sheet.Cells["A5"].PutValue("End");
        sheet.Cells["B5"].PutValue(120);   // Total (end)

        // Add a Waterfall chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Apply custom colors:
        // Point 0 – first total column
        chart.NSeries[0].Points[0].Area.ForegroundColor = Color.Gray;
        // Point 1 – increase column
        chart.NSeries[0].Points[1].Area.ForegroundColor = Color.Green;
        // Point 2 – decrease column
        chart.NSeries[0].Points[2].Area.ForegroundColor = Color.Red;
        // Point 3 – final total column
        chart.NSeries[0].Points[3].Area.ForegroundColor = Color.Gray;

        // Optional: recalculate the chart layout
        chart.Calculate();

        // Save the workbook with the chart
        workbook.Save("WaterfallCustomColors.xlsx", SaveFormat.Xlsx);
    }
}
