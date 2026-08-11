// Title: Add a Chart Legend and Freeze Rows with Aspose.Cells for .NET
// Description: Creates a workbook, inserts sample data, builds a column chart, shows a bottom legend with custom width, height and bold 12‑pt font, freezes rows 1‑20 so the legend stays visible while scrolling, and saves the file as ChartWithLegendAndFrozenRows.xlsx.
// Keywords: Aspose.Cells chart legend | C# freeze panes Aspose.Cells | customize chart legend .NET | FreezePanes example Aspose.Cells | set legend size programmatically | Aspose.Cells column chart tutorial | Excel legend positioning C#
// Common Searches: how to add a legend to a chart using Aspose.Cells C# | freeze rows that contain a chart legend Aspose.Cells | set legend width and height Aspose.Cells .NET | Aspose.Cells FreezePanes to keep header visible | customize chart legend font Aspose.Cells
// Developer Intent: Display a legend on a chart and keep it in view by freezing the rows that include the legend.
// Use Cases: Generate a column chart with a bottom legend that has a fixed size and bold font, then lock the first 20 rows so the legend never scrolls out of sight. | Build an Excel report where the legend acts as a persistent header for large data tables, ensuring readers always see the series identifiers. | Create a dashboard workbook with multiple charts sharing a common legend area that remains static while users navigate through extensive worksheets.
// AI Prompts: Write C# code with Aspose.Cells to add a bottom legend to a column chart, set its width, height, and bold 12‑pt font, then freeze rows 1‑20. | Show an Aspose.Cells example that uses FreezePanes to keep the first 20 rows visible while allowing horizontal scrolling. | Explain how to configure Legend.IsOverlay, Legend.IsAutomaticSize, and font properties for charts in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendFreezeDemo
{
    // Creates a workbook, inserts sample data, builds a column chart, shows a bottom legend with custom width, height and bold 12‑pt font, freezes rows 1‑20 so the legend stays visible while scrolling, and saves the file as ChartWithLegendAndFrozenRows.xlsx.
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

            // Add a column chart (rows 5‑20, columns 0‑8)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is displayed and customize its appearance
            chart.ShowLegend = true;                     // make sure legend is visible
            chart.Legend.Position = LegendPositionType.Bottom; // place legend below the plot area
            chart.Legend.IsOverLay = false;              // legend will not overlap the chart
            chart.Legend.IsAutomaticSize = false;        // allow manual sizing
            chart.Legend.Width = 400;                    // width in pixels
            chart.Legend.Height = 30;                    // height in pixels
            chart.Legend.Font.Size = 12;                 // font size
            chart.Legend.Font.IsBold = true;             // bold font

            // Freeze the rows that contain the legend so it stays visible while scrolling.
            // The legend is positioned at the bottom of the chart (row 20). Freeze rows 1‑20.
            sheet.FreezePanes(20, 0, 20, 0);

            // Save the workbook
            workbook.Save("ChartWithLegendAndFrozenRows.xlsx");
        }
    }
}
