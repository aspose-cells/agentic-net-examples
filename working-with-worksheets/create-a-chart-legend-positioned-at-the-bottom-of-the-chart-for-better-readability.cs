// Title: Aspose.Cells C# – Position Chart Legend at Bottom and Define Custom Size
// Description: Creates a new workbook, fills cells A1:B4 with sample data, adds a column chart, links the data series, moves the legend to the bottom using LegendPositionType.Bottom, disables automatic sizing, sets a custom width and height, and saves the file as ChartWithBottomLegend.xlsx.
// Keywords: Aspose.Cells | C# chart legend position | LegendPositionType.Bottom | custom legend size Aspose.Cells | column chart legend bottom | Excel workbook creation Aspose.Cells | set chart legend dimensions | Aspose.Cells example .NET | chart legend placement | Aspose.Cells legend customization
// Common Searches: Aspose.Cells set legend position bottom | C# chart legend bottom Aspose.Cells example | how to change legend size in Aspose.Cells | LegendPositionType.Bottom usage | customize chart legend dimensions Aspose.Cells
// Developer Intent: The developer needs to move a chart legend to the bottom of the chart and control its width and height.
// Use Cases: Generate a sales‑performance Excel report where the column chart legend is placed below the chart to keep the visual area clear. | Build a dashboard worksheet with a wide column chart and a manually sized bottom legend that fits a predefined layout. | Create a presentation‑ready workbook where the legend is positioned at the bottom to avoid overlapping data series.
// AI Prompts: Show me C# code using Aspose.Cells to set a chart legend to the bottom and specify its width and height. | Provide an Aspose.Cells example that creates a column chart, binds data, moves the legend to the bottom, disables automatic sizing, and saves the workbook. | Explain the effect of the IsAutomaticSize property on chart legends in Aspose.Cells and how to switch between automatic and manual sizing.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendBottomExample
{
    // Creates a new workbook, fills cells A1:B4 with sample data, adds a column chart, links the data series, moves the legend to the bottom using LegendPositionType.Bottom, disables automatic sizing, sets a custom width and height, and saves the file as ChartWithBottomLegend.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the legend at the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Optional: adjust legend size for better appearance
            chart.Legend.IsAutomaticSize = false;
            chart.Legend.Width = 400;
            chart.Legend.Height = 50;

            // Save the workbook (lifecycle: save)
            workbook.Save("ChartWithBottomLegend.xlsx");
        }
    }
}
