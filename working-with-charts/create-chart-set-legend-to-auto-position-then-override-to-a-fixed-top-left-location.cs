// Title: Create a column chart with Aspose.Cells for .NET, set the legend to automatic positioning, then move it to a fixed top‑left location
// AI Prompts: Write C# code using Aspose.Cells to add a column chart, call SetPositionAuto on the legend, then set Legend.Position to NotDocked and assign X/Y coordinates for a top‑left placement. | Show how to customize the size and position of a chart legend after enabling automatic positioning, using Legend.X, Legend.Y, Legend.Width, and Legend.Height in Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart legend to automatic then custom coordinates | How to position a chart legend at a specific X Y location using Aspose.Cells .NET | Move column chart legend to top left corner with Aspose.Cells | Override automatic legend placement in Aspose.Cells chart example
// Tags: Aspose.Cells column chart legend custom position | Legend.SetPositionAuto Aspose.Cells | Legend.Position NotDocked .NET | Chart legend X Y coordinates Aspose.Cells | Aspose.Cells chart legend size adjustment

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendPositionDemo
{
    // The example creates a workbook, fills it with sample data, adds a column chart, sets the legend to automatic positioning, then switches the legend to NotDocked and assigns specific X/Y coordinates, width, and height before saving the file as ChartWithCustomLegendPosition.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
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

            // Access the legend of the chart
            Legend legend = chart.Legend;

            // 1. Set legend to automatic positioning
            legend.SetPositionAuto();

            // 2. Override to a fixed top‑left location
            // Use NotDocked so that X and Y coordinates are respected
            legend.Position = LegendPositionType.NotDocked;

            // X and Y are measured in 1/4000 of the chart area.
            // Adjust these values as needed for the desired placement.
            legend.X = 200;   // Approx. 5% from the left edge
            legend.Y = 200;   // Approx. 5% from the top edge

            // Optionally set size (also in 1/4000 units)
            legend.Width = 800;
            legend.Height = 400;

            // Save the workbook to an XLSX file
            workbook.Save("ChartWithCustomLegendPosition.xlsx");
        }
    }
}
