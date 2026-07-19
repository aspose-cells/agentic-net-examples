// Title: Add a Clickable Hyperlink to a Chart Label with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, populate sample data, generate a column chart, place a label shape inside the chart area, and bind a URL so the label opens a web page when the chart is viewed.
// Keywords: Aspose.Cells | C# | .NET | Excel chart label hyperlink | AddHyperlink | Shapes.AddLabelInChart | clickable chart label | chart annotation link | hyperlink object | interactive Excel chart
// Common Searches: Aspose.Cells add hyperlink to chart label C# | how to make a chart label clickable in Excel using Aspose | C# Aspose.Cells label with URL | add web link to chart annotation Aspose.Cells | interactive chart label Aspose.Cells .NET
// Developer Intent: Create a chart label that functions as a web link.
// Use Cases: Embed a product‑page link in a sales chart label for one‑click access. | Provide a help‑center URL inside a performance chart to guide users to documentation. | Attach a reference report link to a chart title label for quick navigation.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a label into a chart and attaches a URL hyperlink with a custom screen tip. | Show how to add multiple hyperlinks to different labels within the same chart using Aspose.Cells for .NET. | Explain how to change the hyperlink target and display text of an existing chart label after the workbook has been saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to build a workbook, populate sample data, generate a column chart, place a label shape inside the chart area, and bind a URL so the label opens a web page when the chart is viewed.
class AddHyperlinkToChartLabel
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.SetChartDataRange("A1:B4", true);

        // Add a label inside the chart area
        // Parameters are in 1/4000 of the chart area (top, left, height, width)
        Label label = chart.Shapes.AddLabelInChart(1000, 1000, 2000, 4000);
        label.Text = "Visit Aspose";
        label.Font.Color = System.Drawing.Color.Blue;
        label.Font.Size = 12;

        // Add a hyperlink to the label
        // Use Shape.AddHyperlink to create the hyperlink object
        Aspose.Cells.Hyperlink hyperlink = label.AddHyperlink("https://www.aspose.com");
        hyperlink.TextToDisplay = "Aspose Website";
        hyperlink.ScreenTip = "Click to open Aspose";

        // Save the workbook
        workbook.Save("ChartWithLabelHyperlink.xlsx");
    }
}
