// Title: Set per-series legend font colors with transparent background in Aspose.Cells C# chart
// Description: Demonstrates how to create a column chart in Aspose.Cells, then loop through each series to assign a distinct legend font color while keeping the legend entry background transparent and text fill enabled.
// Keywords: Aspose.Cells legend font color | C# chart legend customization | transparent legend background | per series legend color Aspose | Excel chart legend styling | Aspose.Cells Chart API | LegendEntry Font.Color | BackgroundMode.Transparent
// Common Searches: Aspose.Cells change legend font color by series | how to make legend background transparent in Aspose.Cells | set different colors for each legend entry C# | customize chart legend appearance Aspose.Cells | apply series index colors to legend Aspose
// Developer Intent: Apply a specific font color to each legend entry based on its series index while preserving a transparent legend background.
// Use Cases: Generate Excel reports where legend labels match series colors without obscuring underlying cell colors. | Create visually consistent dashboards that overlay charts on colored backgrounds. | Provide a reusable routine for dynamic charts with any number of series, ensuring each legend entry has a distinct font color and no background fill.
// AI Prompts: Write a C# method that receives a Chart object and a Color[] and sets each series' LegendEntry.Font.Color and BackgroundMode to Transparent. | Explain the interaction between LegendEntry.Font.Color, BackgroundMode, and IsTextNoFill in Aspose.Cells. | Show how to handle more series than colors by cycling through a color array when customizing legend fonts.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCustomColor
{
    // Demonstrates how to create a column chart in Aspose.Cells, then loop through each series to assign a distinct legend font color while keeping the legend entry background transparent and text fill enabled.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // Series 1
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Series 3
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.ShowLegend = true;

            // Add the three series to the chart
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.Add("D2:D4", true); // Series 3
            chart.NSeries.CategoryData = "A2:A4";

            // Define a set of colors to apply per series index
            Color[] seriesColors = new Color[]
            {
                Color.Red,
                Color.Green,
                Color.Blue
            };

            // Iterate over each series and customize its legend entry
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                LegendEntry legendEntry = series.LegendEntry;

                // Apply custom font color based on series index
                legendEntry.Font.Color = seriesColors[i % seriesColors.Length];

                // Keep the legend entry background transparent
                legendEntry.BackgroundMode = BackgroundMode.Transparent;

                // Ensure the text itself is not set to "no fill"
                legendEntry.IsTextNoFill = false;
            }

            // Save the workbook
            workbook.Save("ChartWithCustomLegendColors.xlsx");
        }
    }
}
