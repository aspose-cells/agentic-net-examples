// Title: Insert a Bold Arial 12‑pt Label into an Aspose.Cells Chart (C#)
// Description: Shows how to build a workbook, generate a column chart, embed a text label with AddLabelInChart, and apply Arial 12‑pt bold styling using Aspose.Cells for .NET (suitable for developers in the US, Europe, and APAC).
// Keywords: Aspose.Cells AddLabelInChart | chart label Arial 12pt | C# Aspose.Cells label formatting | insert text inside chart .NET | Aspose.Cells chart annotation | label font styling Aspose.Cells | Aspose.Cells chart note example
// Common Searches: How to add a label inside a chart with Aspose.Cells C# | Set chart label font to Arial 12 bold in Aspose.Cells | Aspose.Cells AddLabelInChart positioning coordinates | Change color of chart annotation using Aspose.Cells .NET | Create custom note inside a chart with Aspose.Cells
// Developer Intent: Place a text annotation within a chart area and format it with Arial, 12‑pt, bold font via the Aspose.Cells API.
// Use Cases: Add an explanatory note to a financial column chart without using external shapes. | Highlight a key metric inside a sales chart with a styled label for presentations. | Create a custom legend or disclaimer directly inside a chart for automated reporting.
// AI Prompts: Generate C# code that adds a centered Times New Roman 14‑pt italic label to an Aspose.Cells chart. | Show how to insert multiple labels at different positions in a chart using AddLabelInChart with varied font settings. | Explain how to bind a chart label's text to a worksheet cell so the label updates automatically when the cell changes.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLabelInChartDemo
{
    // Shows how to build a workbook, generate a column chart, embed a text label with AddLabelInChart, and apply Arial 12‑pt bold styling using Aspose.Cells for .NET (suitable for developers in the US, Europe, and APAC).
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a label inside the chart using AddLabelInChart (top, left, height, width)
            // Values are in 1/4000 of the chart area
            Label chartLabel = chart.Shapes.AddLabelInChart(1000, 1000, 2000, 4000);
            chartLabel.Text = "Chart Label";

            // Set the label's font to Arial, size 12, bold
            chartLabel.Font.Name = "Arial";
            chartLabel.Font.Size = 12;
            chartLabel.Font.IsBold = true;
            chartLabel.Font.Color = Color.Black;

            // Save the workbook
            workbook.Save("LabelInChartDemo.xlsx");
        }
    }
}
