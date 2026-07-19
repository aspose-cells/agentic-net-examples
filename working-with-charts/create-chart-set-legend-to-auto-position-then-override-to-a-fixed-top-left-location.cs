// Title: Aspose.Cells .NET: Create a column chart, set legend to auto, then move it to a fixed top‑left position
// Description: This example shows how to generate an Excel workbook with a column chart, apply automatic legend placement using SetPositionAuto(), switch the legend to NotDocked, and assign XPixel/YPixel coordinates (e.g., 100,100) for a precise top‑left location. Optional width and height settings are also demonstrated before saving the file as ChartLegendAutoThenFixed.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel chart legend position | SetPositionAuto | LegendPositionType.NotDocked | XPixels YPixels | fixed legend coordinates | column chart example | programmatic legend sizing
// Common Searches: Aspose.Cells set legend auto then fixed position | C# chart legend NotDocked XPixel YPixel | how to move Excel chart legend with Aspose.Cells | auto legend placement Aspose.Cells .NET | custom legend size Excel chart C#
// Developer Intent: Create a chart, let the legend auto‑place, then override it with exact pixel coordinates.
// Use Cases: Design a sales dashboard where the legend must appear at a consistent spot across multiple charts. | Produce a financial report that prevents legend overlap by fixing its position after initial auto‑placement. | Automate layout adjustments for a batch of charts, ensuring uniform legend size and location in each worksheet.
// AI Prompts: Generate C# code with Aspose.Cells to add a line chart, set the legend to automatic, then relocate it to (50,30) using NotDocked. | Explain the effect of LegendPositionType.NotDocked on XPixel and YPixel values in Aspose.Cells charts. | Provide a step‑by‑step guide for customizing legend dimensions and coordinates for various chart types in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to generate an Excel workbook with a column chart, apply automatic legend placement using SetPositionAuto(), switch the legend to NotDocked, and assign XPixel/YPixel coordinates (e.g., 100,100) for a precise top‑left location. Optional width and height settings are also demonstrated before saving the file as ChartLegendAutoThenFixed.xlsx.
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

        // 1. Set legend to automatic positioning
        chart.Legend.SetPositionAuto();

        // 2. Override legend to a fixed top‑left location
        //    Use NotDocked so manual coordinates are respected
        chart.Legend.Position = LegendPositionType.NotDocked;
        // Position in pixels relative to the chart area
        chart.Legend.XPixel = 100; // distance from the left edge
        chart.Legend.YPixel = 100; // distance from the top edge

        // Optional: set a fixed size for the legend
        chart.Legend.WidthPixel = 150;
        chart.Legend.HeightPixel = 80;

        // Save the workbook with the chart
        workbook.Save("ChartLegendAutoThenFixed.xlsx");
    }
}
