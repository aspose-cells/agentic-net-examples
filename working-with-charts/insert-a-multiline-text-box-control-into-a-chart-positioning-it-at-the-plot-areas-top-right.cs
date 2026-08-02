// Title: Aspose.Cells for .NET – Insert a Multiline TextBox in a Chart’s Plot Area (Top‑Right)
// Description: This example creates a workbook, adds sample data, builds a column chart, and uses chart.Shapes.AddTextBoxInChart to place a multiline TextBox at the plot area’s top‑right corner. The box contains line‑breaks, custom font color and size, overflow handling, and the file is saved as MultilineTextboxInChart.xlsx.
// Keywords: Aspose.Cells AddTextBoxInChart | multiline textbox chart .NET | position textbox top right chart | C# chart shape Aspose.Cells | textbox overflow Aspose.Cells
// Common Searches: how to add a multiline textbox to a chart using Aspose.Cells | place textbox at top right of chart plot area C# | Aspose.Cells AddTextBoxInChart coordinate system | set font color and size for chart textbox Aspose.Cells | enable text overflow in chart textbox .NET
// Developer Intent: Add a multiline TextBox shape to a chart and locate it at the plot area’s top‑right corner.
// Use Cases: Add explanatory notes beside a chart without covering data series. | Create a multi‑line description or legend inside the chart area. | Display dynamic comments that can expand when the workbook is regenerated.
// AI Prompts: Write C# code with Aspose.Cells to insert a multiline textbox at a specific chart position and style its font. | Explain the 1/4000 unit coordinate system used by AddTextBoxInChart and how to compute offsets for different chart sizes. | Show how to bind the textbox text to worksheet cells so the content updates automatically.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds sample data, builds a column chart, and uses chart.Shapes.AddTextBoxInChart to place a multiline TextBox at the plot area’s top‑right corner. The box contains line‑breaks, custom font color and size, overflow handling, and the file is saved as MultilineTextboxInChart.xlsx.
    class InsertMultilineTextboxInChart
    {
        public static void Run()
        {
            try
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
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Add a multiline textbox to the chart positioned at the plot area's top‑right.
                // Units are 1/4000 of the chart area.
                int top = 0;          // vertical offset from the top
                int left = 3500;      // horizontal offset near the right edge
                int height = 500;     // height of the textbox
                int width = 500;      // width of the textbox

                TextBox textBox = chart.Shapes.AddTextBoxInChart(top, left, height, width);
                textBox.Text = "First line\nSecond line\nThird line"; // multiline content
                textBox.Font.Color = Color.DarkBlue;
                textBox.Font.Size = 12;
                textBox.TextBoxOptions.AllowTextToOverflow = true; // ensure all text is visible

                // Save the workbook
                workbook.Save("MultilineTextboxInChart.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            InsertMultilineTextboxInChart.Run();
        }
    }
}
