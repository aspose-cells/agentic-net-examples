// Title: C# – Add a multiline TextBox to a chart’s plot area (top‑right) with Aspose.Cells
// Description: This .NET example creates a workbook, fills it with sample data, builds a column chart, calculates offsets, inserts a TextBox shape inside the chart’s plot area at the upper‑right corner, populates it with three lines of text, enables overflow, and saves the result as MultilineTextboxInChart.xlsx.
// Keywords: Aspose.Cells C# chart textbox | multiline textbox Aspose.Cells | position textbox top right chart | chart shape placement Aspose.Cells | AllowTextToOverflow Aspose.Cells | AddTextBoxInChart method | Excel chart annotation C#
// Common Searches: how to add a multiline textbox to a chart using Aspose.Cells .NET | position textbox at the top right of a chart plot area Aspose.Cells | C# Aspose.Cells insert textbox inside chart | enable AllowTextToOverflow for chart textbox Aspose.Cells | Aspose.Cells AddTextBoxInChart example
// Developer Intent: Insert a multi‑line TextBox shape into a chart’s plot area and align it to the upper‑right corner.
// Use Cases: Add explanatory notes or comments inside a chart without covering data series. | Create multi‑line annotations that stay anchored to the chart’s top‑right corner. | Allow the textbox to expand automatically when its content exceeds the defined size.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a multiline textbox into a chart and aligns it to the top‑right of the plot area. | Show how to enable AllowTextToOverflow for a chart textbox using Aspose.Cells for .NET. | Explain the calculation of left and top offsets needed to position a textbox at the right edge of a chart’s plot area.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// This .NET example creates a workbook, fills it with sample data, builds a column chart, calculates offsets, inserts a TextBox shape inside the chart’s plot area at the upper‑right corner, populates it with three lines of text, enables overflow, and saves the result as MultilineTextboxInChart.xlsx.
class InsertMultilineTextboxInChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Define textbox size (in 1/4000 of chart area)
        int txtHeight = 800;   // height
        int txtWidth  = 1200;  // width

        // Position the textbox at the top‑right corner of the plot area.
        // Top offset = 0 (very top). Left offset = chart area width - textbox width.
        // Approximate chart area width as 4000 units, so left = 4000 - txtWidth.
        int leftOffset = 4000 - txtWidth; // right‑most position
        int topOffset  = 0;                // top‑most position

        // Add the textbox to the chart
        TextBox txtBox = chart.Shapes.AddTextBoxInChart(topOffset, leftOffset, txtHeight, txtWidth);

        // Set multiline text (use newline characters)
        txtBox.Text = "Line 1 of text\nLine 2 of text\nLine 3 of text";

        // Allow the text to overflow if needed (optional)
        txtBox.TextBoxOptions.AllowTextToOverflow = true;

        // Save the workbook
        workbook.Save("MultilineTextboxInChart.xlsx");
    }
}
