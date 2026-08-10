// Title: Aspose.Cells .NET: Show Value & Percentage in Stacked Bar Chart Labels & Resize Shapes
// Description: Creates a workbook, adds sample data, builds a stacked bar chart, enables data labels to display both cell values and percentages, disables auto‑fit and sets a fixed label width, resizes the chart to 600 × 400 px, inserts a rectangle shape inside the chart area with LightBlue fill and DarkBlue border, resizes the shape to 150 × 80 px, and saves the file.
// Keywords: Aspose.Cells | C# | stacked bar chart | data labels value percentage | resize chart | shape resizing | disable auto‑fit | chart object size | add rectangle shape | Excel automation
// Common Searches: Aspose.Cells show value and percentage in chart data labels | Resize chart object Aspose.Cells .NET | Disable auto fit for chart data labels Aspose.Cells | Add rectangle shape inside chart Aspose.Cells | Set fixed width for data label shapes
// Developer Intent: Create a stacked bar chart that displays both numeric values and their percentage contributions, and control the dimensions of the chart and an inner rectangle shape.
// Use Cases: Sales or budget reports that need absolute numbers and share percentages on each bar segment. | Standardized Excel templates where chart size must stay constant across pages. | Custom visual cues inside a chart, such as a colored call‑out box, with precise pixel dimensions. | Automated generation of presentation‑ready charts for dashboards.
// AI Prompts: Generate C# code with Aspose.Cells to add a stacked bar chart, enable value and percentage data labels, set a fixed label width, and resize the chart to 600 × 400 px. | Provide a snippet that inserts a rectangle shape into a chart, applies LightBlue fill, DarkBlue border, and resizes it to 150 × 80 px using Aspose.Cells. | Explain how to turn off auto‑fit for chart data label shapes and assign a custom width for each point in a stacked bar chart with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a stacked bar chart, enables data labels to display both cell values and percentages, disables auto‑fit and sets a fixed label width, resizes the chart to 600 × 400 px, inserts a rectangle shape inside the chart area with LightBlue fill and DarkBlue border, resizes the shape to 150 × 80 px, and saves the file.
    public class StackedBarChartDataLabelsAndShapeResize
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked bar chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(40);

            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(30);
            sheet.Cells["C4"].PutValue(10);

            sheet.Cells["D1"].PutValue("Product C");
            sheet.Cells["D2"].PutValue(10);
            sheet.Cells["D3"].PutValue(40);
            sheet.Cells["D4"].PutValue(20);

            // Add a stacked bar chart
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the chart (including all series)
            chart.SetChartDataRange("A1:D4", true);
            chart.Calculate();

            // Enable data labels for each series and show both value and percentage
            foreach (Series ser in chart.NSeries)
            {
                ser.DataLabels.ShowValue = true;
                ser.DataLabels.ShowPercentage = true;

                // Adjust label shape to avoid auto‑fit (fixed width)
                foreach (ChartPoint pt in ser.Points)
                {
                    pt.DataLabels.IsResizeShapeToFitText = false; // disable auto‑fit
                    pt.DataLabels.Width = 80;                    // fixed width in pixels
                }
            }

            // Resize the chart shape directly (avoid obsolete properties)
            chart.ChartObject.Width = 600;   // width in pixels
            chart.ChartObject.Height = 400;  // height in pixels

            // Add an additional shape inside the chart area and resize it
            Shape rect = chart.Shapes.AddShapeInChart(
                MsoDrawingType.Rectangle,
                PlacementType.Move,
                1000,   // left  (1/4000 of chart width)
                1000,   // top   (1/4000 of chart height)
                3000,   // right
                2000);  // bottom

            rect.Fill.SolidFill.Color = Color.LightBlue;
            rect.Line.SolidFill.Color = Color.DarkBlue;
            rect.Width = 150;   // width in pixels
            rect.Height = 80;   // height in pixels

            // Save the workbook
            workbook.Save("StackedBarChart_With_Labels_And_ResizedShapes.xlsx");
        }
    }
}
