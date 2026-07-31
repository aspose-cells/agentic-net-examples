// Title: C# – Add a Legend to a Column Chart and Freeze Its Rows Using Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example creates a workbook, populates sample data, inserts a column chart, displays and customizes the legend (bottom position, fixed width/height, bold font), then freezes the first 20 rows so the chart and legend stay visible while scrolling, and saves the file as ChartWithLegendAndFrozenRows.xlsx.
// Keywords: Aspose.Cells | C# | .NET | add chart legend | column chart | freeze panes | FreezePanes | chart legend position | legend size | custom legend font | Excel automation | chart formatting | worksheet freeze rows
// Common Searches: Aspose.Cells add legend to chart | How to freeze rows with a chart in Aspose.Cells | Set legend width and height Aspose.Cells C# | FreezePanes example Aspose.Cells .NET | Customize chart legend position Aspose.Cells
// Developer Intent: Add a visible legend to a column chart and lock the rows that contain the chart so they remain in view.
// Use Cases: Generate a sales report where the column chart legend must stay on screen while users scroll through data rows. | Build an Excel dashboard that keeps the chart legend fixed by freezing the top rows after inserting and styling the chart.
// AI Prompts: Write C# code with Aspose.Cells to add a bottom‑positioned legend to a column chart, set its width, height, and bold font, then freeze the first 20 rows so the chart and legend stay visible. | Show an Aspose.Cells .NET example that inserts a column chart, customizes the legend appearance, and uses FreezePanes to lock the rows containing the chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendFreezeDemo
{
    // This Aspose.Cells for .NET example creates a workbook, populates sample data, inserts a column chart, displays and customizes the legend (bottom position, fixed width/height, bold font), then freezes the first 20 rows so the chart and legend stay visible while scrolling, and saves the file as ChartWithLegendAndFrozenRows.xlsx.
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

            // Add a column chart (positioned from row 5 to row 20)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is displayed and customize its appearance
            chart.ShowLegend = true;                                 // Show the legend
            chart.Legend.Position = LegendPositionType.Bottom;       // Place legend at the bottom of the chart
            chart.Legend.IsAutomaticSize = false;                    // Disable automatic sizing
            chart.Legend.Width = 400;                                // Set legend width (pixels)
            chart.Legend.Height = 50;                                // Set legend height (pixels)
            chart.Legend.Font.Size = 12;                             // Optional: adjust font size
            chart.Legend.Font.IsBold = true;                         // Optional: make font bold

            // Freeze the rows that contain the chart (including the legend) so they stay visible while scrolling
            // Freeze the top 20 rows (rows 0‑19) and no columns
            sheet.FreezePanes(20, 0, 20, 0);

            // Save the workbook
            workbook.Save("ChartWithLegendAndFrozenRows.xlsx");
        }
    }
}
