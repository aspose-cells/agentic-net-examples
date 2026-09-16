// Title: Add a multiline TextBox to the top‑right corner of a chart’s plot area using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a multiline TextBox shape into the plot area of a column chart at the top‑right position with Aspose.Cells for C#. | Calculate the plot area coordinates of a chart and add a centered multiline textbox inside it using the Aspose.Cells .NET API. | Set horizontal and vertical center alignment for a chart textbox and optionally apply fill and border colors in C#.
// Common Searches: Aspose.Cells C# add multiline textbox to chart plot area | position textbox at top right of chart using Aspose.Cells .NET | how to place a shape inside a chart with Aspose.Cells | center align text inside chart textbox Aspose.Cells C# | save workbook with chart and textbox Aspose.Cells example
// Tags: add multiline textbox to chart Aspose.Cells | chart plot area shape positioning .NET | Aspose.Cells textbox alignment | insert shape into column chart C# | Aspose.Cells chart formatting example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, populating sample data, adding a column chart, computing the top‑right coordinates of the chart’s plot area, inserting a multiline TextBox shape inside the chart, setting its text and center alignment, optionally configuring fill and border styling, and saving the workbook with Aspose.Cells for .NET.
class InsertMultilineTextboxInChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart data source (values)
            chart.NSeries.Add("B2:B4", true);
            // Category (X‑axis) data is taken automatically from the first column of the range

            // Define textbox size (points)
            int textboxWidth = 120;
            int textboxHeight = 60;

            // Calculate position: top‑right corner of the plot area
            int leftPosition = chart.PlotArea.X + chart.PlotArea.Width - textboxWidth;
            int topPosition = chart.PlotArea.Y;

            // Add a multiline textbox inside the chart
            TextBox textbox = chart.Shapes.AddTextBoxInChart(leftPosition, topPosition, textboxWidth, textboxHeight);

            // Set multiline text
            textbox.Text = "First line\nSecond line\nThird line";

            // Optional formatting
            textbox.TextHorizontalAlignment = TextAlignmentType.Center;
            textbox.TextVerticalAlignment = TextAlignmentType.Center;
            // Background fill and border formatting can be set if needed, e.g.:
            // textbox.Fill.ForeColor = Color.LightYellow;
            // textbox.Line.ForeColor = Color.DarkGray;
            // textbox.Line.Weight = 1;

            // Save the workbook
            string outputPath = "ChartWithMultilineTextbox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
