// Title: How to add a multiline textbox to the top‑right corner of a chart’s plot area using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a textbox shape, sets its width and height, positions it at the chart’s plot‑area top‑right corner, and fills it with multiline text using Aspose.Cells. | Write a C# snippet that calculates the left and top coordinates from Chart.PlotArea and adds a centered multiline textbox to a column chart in an Aspose.Cells workbook. | Provide C# instructions to add a resizable textbox shape to a chart, align its text both horizontally and vertically, and save the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# add textbox to chart plot area top right | place multiline text box in Aspose.Cells chart using C# | calculate chart plot area coordinates for shape positioning Aspose.Cells | insert a textbox shape with line breaks into a column chart Aspose.Cells | Aspose.Cells set textbox alignment and size in chart
// Tags: Aspose.Cells add textbox to chart | C# chart plot area shape positioning | multiline textbox Aspose.Cells | textbox dimensions alignment Aspose.Cells | column chart textbox Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, then creates a 120 × 50‑point textbox shape, positions it at the chart’s plot‑area top‑right corner, sets multiline text with line breaks, centers the text horizontally and vertically, unlocks the shape for editing, and saves the file as ChartWithMultilineTextbox.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define textbox dimensions (in points)
            double tbWidth = 120;   // width of the textbox
            double tbHeight = 50;   // height of the textbox

            // Position the textbox at the plot area's top‑right corner
            double left = chart.PlotArea.X + chart.PlotArea.Width - tbWidth;
            double top = chart.PlotArea.Y;

            // Add a textbox shape (initially at (0,0)) and then set its position
            Shape txtBox = chart.Shapes.AddTextBox(0, 0, 0, 0, (int)tbWidth, (int)tbHeight);
            txtBox.Left = (int)left;   // Shape.Left expects an int
            txtBox.Top = (int)top;     // Shape.Top expects an int

            // Set textbox content and formatting
            txtBox.Text = "First line\r\nSecond line\r\nThird line";
            txtBox.TextHorizontalAlignment = TextAlignmentType.Center;
            txtBox.TextVerticalAlignment = TextAlignmentType.Center;
            txtBox.IsLocked = false; // allow editing if needed

            // Save the workbook
            wb.Save("ChartWithMultilineTextbox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
