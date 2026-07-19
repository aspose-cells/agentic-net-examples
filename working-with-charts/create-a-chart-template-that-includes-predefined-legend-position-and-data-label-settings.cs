// Title: Aspose.Cells .NET – Create a Chart Template with Fixed Legend Position and Styled Data Labels
// Description: Demonstrates how to generate an Excel workbook, populate a simple data range, add a column chart, set the legend to the bottom without overlay, enable data labels, and format those labels with a 10‑point bold blue font, then save the file as a reusable chart template (ChartTemplate.xlsx).
// Keywords: Aspose.Cells chart legend position | Aspose.Cells data label formatting | C# chart template Aspose.Cells | Excel chart styling .NET | disable legend overlay Aspose.Cells | column chart data labels bold blue
// Common Searches: set legend position bottom Aspose.Cells chart | disable legend overlay in Aspose.Cells | show and style data labels for column chart C# | save chart as template Aspose.Cells | apply consistent chart formatting across workbooks
// Developer Intent: Create a reusable chart template that enforces a bottom legend without overlay and applies bold blue data label styling.
// Use Cases: Standardize chart appearance for automated reporting dashboards. | Batch‑process Excel exports where every column chart must display values in a specific font and color. | Provide a base workbook for users to add data while preserving predefined legend and label settings.
// AI Prompts: Generate C# code using Aspose.Cells to add a column chart with the legend at the bottom, overlay disabled, and data labels shown in bold blue 10‑point font. | Show how to apply the same legend and data label configuration to an existing chart in a workbook and save it as a reusable template. | Explain how to reuse ChartTemplate.xlsx as a starter file for other workbooks while keeping the predefined chart formatting intact.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to generate an Excel workbook, populate a simple data range, add a column chart, set the legend to the bottom without overlay, enable data labels, and format those labels with a 10‑point bold blue font, then save the file as a reusable chart template (ChartTemplate.xlsx).
class ChartTemplateDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Add a column chart to the worksheet (uses ChartCollection.Add rule)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B5", true);

        // Predefine legend position and overlay behavior (uses Legend.Position and Legend.IsOverLay)
        chart.Legend.Position = LegendPositionType.Bottom;
        chart.Legend.IsOverLay = false;

        // Enable data labels to show values and customize their appearance
        chart.NSeries[0].DataLabels.ShowValue = true;
        chart.NSeries[0].DataLabels.Font.Size = 10;
        chart.NSeries[0].DataLabels.Font.IsBold = true;
        chart.NSeries[0].DataLabels.Font.Color = Color.Blue;

        // Recalculate the chart layout after modifications
        chart.Calculate();

        // Save the workbook with the chart template
        workbook.Save("ChartTemplate.xlsx", SaveFormat.Xlsx);
    }
}
