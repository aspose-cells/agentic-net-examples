// Title: Add a multiline TextBox to a chart’s plot area (top‑right) with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, and uses Chart.Shapes.AddTextBoxInChart to insert a multiline TextBox at the plot area’s top‑right corner (coordinates 0,3000) with custom height, width, and AllowTextToOverflow enabled, then saves the file.
// Keywords: Aspose.Cells | C# chart textbox | AddTextBoxInChart | multiline textbox Aspose.Cells | chart shape positioning | top right chart textbox | text overflow chart | Aspose.Cells example | Excel chart annotation | plot area coordinates
// Common Searches: how to add a textbox to a chart using Aspose.Cells | multiline textbox in Aspose.Cells chart | position chart textbox top right | allow text overflow in chart textbox Aspose.Cells | C# example AddTextBoxInChart | set chart shape coordinates Aspose.Cells | Aspose.Cells chart annotation tutorial
// Developer Intent: Insert a multiline TextBox shape into a chart and position it at the plot area’s top‑right corner.
// Use Cases: Add explanatory notes to a column chart with line breaks for clarity. | Create a dashboard element that may exceed its box size, using AllowTextToOverflow to prevent clipping. | Place a dynamic title or timestamp at the chart’s top‑right without affecting the legend.
// AI Prompts: Write C# code that uses Aspose.Cells to add a multiline TextBox to a chart, set line‑break text, enable overflow, and position it at the chart’s top‑right. | Explain how to calculate the left coordinate for AddTextBoxInChart to align a textbox with the right edge of the plot area. | Show how to modify the size, position, and text properties of an existing TextBoxInChart in Aspose.Cells after it has been added.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a column chart, and uses Chart.Shapes.AddTextBoxInChart to insert a multiline TextBox at the plot area’s top‑right corner (coordinates 0,3000) with custom height, width, and AllowTextToOverflow enabled, then saves the file.
class InsertMultilineTextboxInChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["A5"].PutValue("D");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["B5"].PutValue(40);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Bind data to the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Insert a multiline TextBox into the chart.
        // Units are 1/4000 of the chart area.
        // Position it at the top‑right corner of the plot area.
        // Top = 0 (top edge), Left = 3000 (near right edge), Height = 500, Width = 800.
        TextBox textBox = chart.Shapes.AddTextBoxInChart(0, 3000, 500, 800);

        // Set multiline text using line breaks
        textBox.Text = "First line\nSecond line\nThird line";

        // Allow the text to overflow if needed (optional)
        textBox.TextBoxOptions.AllowTextToOverflow = true;

        // Save the workbook
        workbook.Save("MultilineTextboxInChart.xlsx");
    }
}
