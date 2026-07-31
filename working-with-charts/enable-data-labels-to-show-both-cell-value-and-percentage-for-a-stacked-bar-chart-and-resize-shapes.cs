// Title: Show Value & Percentage in Stacked Bar Chart Labels and Resize Chart/Shapes – Aspose.Cells for .NET
// Description: Creates a workbook, adds a stacked bar chart, configures each series to display both the cell value and its percentage in data labels, locks label dimensions, resizes the chart object, and inserts a rectangle shape positioned by percentage coordinates.
// Keywords: Aspose.Cells stacked bar chart | data labels value and percentage | C# chart resizing | fixed label size Aspose.Cells | add shape by scale Aspose.Cells | .NET Excel chart example
// Common Searches: how to show value and percentage in Aspose.Cells chart labels | resize chart object C# Aspose.Cells | prevent data label auto‑size Aspose.Cells | insert rectangle shape in chart using percentage coordinates | Aspose.Cells stacked bar chart tutorial
// Developer Intent: Generate a stacked bar chart that presents both absolute values and percentages in its labels, enforce uniform label dimensions, adjust the chart size, and embed a styled annotation shape.
// Use Cases: Financial statements where each segment needs its amount and share displayed. | Executive dashboards that require consistent label sizing for readability. | Highlighting a specific range in a chart with a colored annotation rectangle.
// AI Prompts: Write C# code with Aspose.Cells to add a stacked bar chart whose data labels show both the cell value and percentage, and set a fixed width and height for the labels. | Provide an example that resizes the ChartObject and places a rectangle shape inside the chart using scale (percentage) coordinates, then apply custom fill and line colors. | Explain how to disable automatic resizing of data label shapes in Aspose.Cells and why fixed dimensions are useful for report consistency.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsStackedBarDataLabels
{
    // Creates a workbook, adds a stacked bar chart, configures each series to display both the cell value and its percentage in data labels, locks label dimensions, resizes the chart object, and inserts a rectangle shape positioned by percentage coordinates.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series 1 values
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Series 2 values
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(110);
            sheet.Cells["C4"].PutValue(130);
            sheet.Cells["C5"].PutValue(170);

            // Add a stacked bar chart
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";     // Categories

            // Enable data labels to show both value and percentage for each series
            foreach (Series ser in chart.NSeries)
            {
                ser.DataLabels.ShowValue = true;          // Show cell value
                ser.DataLabels.ShowPercentage = true;    // Show percentage

                // Prevent automatic resizing of the label shape so we can set a fixed size
                ser.DataLabels.IsResizeShapeToFitText = false;
                ser.DataLabels.WidthPixel = 80;   // Fixed width in pixels
                ser.DataLabels.HeightPixel = 30;  // Fixed height in pixels
            }

            // Resize the whole chart shape (ChartObject) to a specific size
            chart.ChartObject.Width = 600;   // Width in pixels
            chart.ChartObject.Height = 350;  // Height in pixels

            // Add a rectangle shape inside the chart area using percentage coordinates
            Shape rect = chart.Shapes.AddShapeInChartByScale(
                MsoDrawingType.Rectangle,
                PlacementType.Move,
                0.10,   // left  = 10% of chart width
                0.10,   // top   = 10% of chart height
                0.30,   // right = 30% of chart width
                0.25);  // bottom= 25% of chart height

            // Format the added shape
            rect.Fill.SolidFill.Color = Color.LightBlue;
            rect.Line.SolidFill.Color = Color.DarkBlue;
            rect.Text = "Custom Annotation";

            // Save the workbook
            workbook.Save("StackedBar_With_Value_And_Percentage.xlsx");
        }
    }
}
