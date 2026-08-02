// Title: Add a Hyperlink to a Chart Label with Aspose.Cells (C#)
// Description: Creates a workbook, builds a column chart, inserts a formatted label inside the chart area, attaches a hyperlink that opens a web page when the label is clicked, and saves the workbook.
// Keywords: Aspose.Cells chart label hyperlink | C# add hyperlink to chart shape | Excel chart clickable label | Aspose.Cells label in chart | ChartLabel.AddHyperlink example | Aspose.Cells C# hyperlink demo | Excel hyperlink label Aspose | Add URL to chart label C#
// Common Searches: Aspose.Cells add hyperlink to chart label C# | How to make a chart label clickable in Excel using Aspose | C# chart label hyperlink Aspose.Cells | Add URL to chart label Aspose | Set screen tip for chart label hyperlink Aspose.Cells
// Developer Intent: Attach a web URL to a label shape inside an Excel chart using Aspose.Cells for C#.
// Use Cases: Embed a label in a sales chart that links directly to the product page. | Provide a quick‑access link from a dashboard chart to external documentation. | Show a tooltip (screen tip) on a chart label that explains the destination before the user clicks.
// AI Prompts: Generate C# code with Aspose.Cells that adds a label to a chart and attaches a hyperlink to https://example.com, including display text and a screen tip. | Explain how to configure the chart label hyperlink to open in a new browser tab when using Aspose.Cells. | Show how to read, modify, or remove the hyperlink of an existing chart label after the workbook has been saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLabelHyperlinkDemo
{
    // Creates a workbook, builds a column chart, inserts a formatted label inside the chart area, attaches a hyperlink that opens a web page when the label is clicked, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.SetChartDataRange("A1:B4", true);

            // Add a label inside the chart area
            // Parameters are in 1/4000 of the chart area (top, left, height, width)
            Label chartLabel = chart.Shapes.AddLabelInChart(1000, 1000, 2000, 3000);
            chartLabel.Text = "Visit Aspose";
            chartLabel.Font.Color = System.Drawing.Color.Blue;
            chartLabel.Font.Size = 12;

            // Add a hyperlink to the label shape
            // Use Shape.AddHyperlink which returns a Hyperlink object
            Aspose.Cells.Hyperlink link = chartLabel.AddHyperlink("https://www.aspose.com");
            // Optionally set display text and screen tip
            link.TextToDisplay = "Aspose Website";
            link.ScreenTip = "Click to open Aspose site";

            // Save the workbook
            workbook.Save("ChartLabelWithHyperlink.xlsx");
        }
    }
}
