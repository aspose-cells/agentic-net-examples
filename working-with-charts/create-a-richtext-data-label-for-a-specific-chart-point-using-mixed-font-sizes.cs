// Title: Aspose.Cells .NET – Create a Rich‑Text Data Label with Mixed Font Sizes for a Chart Point
// Description: Demonstrates how to add a column chart to a workbook, enable a data label for the first point, set custom text, and use the Characters(startIndex, length) API to apply a larger bold blue font to the prefix and a smaller dark‑red font to the numeric value before saving the file.
// Keywords: Aspose.Cells | C# | .NET | chart data label | rich text label | mixed font size | font color formatting | Characters method | Excel column chart | custom label styling
// Common Searches: Aspose.Cells format part of a chart label | rich text data label C# Aspose.Cells | apply different fonts to chart point label | Characters method chart label Aspose.Cells .NET | custom font size and color for Excel chart data label
// Developer Intent: Add a data label to a specific chart point and style different text fragments with separate font sizes and colors using Aspose.Cells for .NET.
// Use Cases: Display a descriptive prefix in a prominent style while keeping the numeric value subtle on a column chart. | Highlight key performance indicators in Excel reports by applying distinct formatting to parts of a data label. | Generate automated spreadsheets where chart labels need brand‑consistent colors and typography.
// AI Prompts: Write C# code with Aspose.Cells that creates a chart point label where the word "Total" is 16 pt bold green and the value is 10 pt regular gray. | Explain the role of the Characters(startIndex, length) function for rich‑text formatting of chart data labels in Aspose.Cells. | Provide a variant that adds a second line to the label, aligns it center, and uses italic styling for the new line.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRichTextDataLabelDemo
{
    // Demonstrates how to add a column chart to a workbook, enable a data label for the first point, set custom text, and use the Characters(startIndex, length) API to apply a larger bold blue font to the prefix and a smaller dark‑red font to the numeric value before saving the file.
    public class Program
    {
        public static void Main()
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

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first point of the first series
            ChartPoint point = chart.NSeries[0].Points[0];

            // Enable the data label for this point
            point.DataLabels.ShowValue = true;

            // Set a custom text for the data label
            // Example: "Val:10"
            point.DataLabels.Text = "Val:10";

            // Apply rich‑text formatting:
            // Make "Val:" larger (size 14) and bold,
            // keep the numeric part smaller (size 10).
            // Characters(startIndex, length) works on zero‑based index.
            // "Val:" -> start 0, length 4
            point.DataLabels.Characters(0, 4).Font.Size = 14;
            point.DataLabels.Characters(0, 4).Font.IsBold = true;
            point.DataLabels.Characters(0, 4).Font.Color = Color.Blue;

            // "10" -> start 4, length 2
            point.DataLabels.Characters(4, 2).Font.Size = 10;
            point.DataLabels.Characters(4, 2).Font.Color = Color.DarkRed;

            // Apply the font settings to all child nodes of the data label
            point.DataLabels.ApplyFont();

            // Save the workbook
            workbook.Save("RichTextDataLabelDemo.xlsx");
        }
    }
}
