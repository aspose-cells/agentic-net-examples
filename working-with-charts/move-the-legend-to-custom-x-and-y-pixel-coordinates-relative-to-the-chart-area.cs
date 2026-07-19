// Title: Aspose.Cells for .NET – Set Chart Legend Position Using XPixel and YPixel
// Description: Creates a workbook, adds a column chart, and moves the legend to a free‑floating spot by setting Legend.Position to NotDocked and specifying XPixel/YPixel offsets relative to the chart area. Optionally disables overlay before saving.
// Keywords: Aspose.Cells legend position | C# chart legend pixel coordinates | NotDocked legend Aspose.Cells | custom legend XPixel YPixel | Excel chart layout .NET
// Common Searches: Aspose.Cells set legend XPixel | C# move chart legend to pixel coordinates | NotDocked legend example Aspose.Cells | custom legend placement in Excel chart | disable legend overlay Aspose.Cells
// Developer Intent: Place a chart legend at exact pixel offsets inside the chart area using Aspose.Cells for .NET.
// Use Cases: Prevent data overlap by positioning the legend away from chart series. | Align legends with other visual elements in automated reports. | Standardize legend placement across multiple generated charts.
// AI Prompts: Write C# code that positions a pie chart legend at XPixel 200 and YPixel 80 using Aspose.Cells. | Explain the effect of the IsOverLay property when a legend is moved to a custom location. | Show how to set different pixel positions for several chart legends in the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCustomPosition
{
    // Creates a workbook, adds a column chart, and moves the legend to a free‑floating spot by setting Legend.Position to NotDocked and specifying XPixel/YPixel offsets relative to the chart area. Optionally disables overlay before saving.
    class Program
    {
        static void Main()
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the legend and set it to a free‑floating position
            Legend legend = chart.Legend;
            legend.Position = LegendPositionType.NotDocked;   // Allows manual positioning
            legend.XPixel = 150;   // Custom X coordinate in pixels relative to the chart area
            legend.YPixel = 50;    // Custom Y coordinate in pixels relative to the chart area

            // Optionally, disable overlay so the legend does not cover chart data
            legend.IsOverLay = false;

            // Save the workbook
            workbook.Save("ChartWithCustomLegendPosition.xlsx");
        }
    }
}
