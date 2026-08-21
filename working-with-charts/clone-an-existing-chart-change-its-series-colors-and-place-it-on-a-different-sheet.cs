// Title: Clone a Chart, Recolor Its Series, and Move It to Another Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart, duplicate the worksheet (which copies the chart), change the cloned chart's series colors using a built‑in palette, reposition the chart on the new sheet, and save the file.
// Keywords: Aspose.Cells chart clone C# | change chart series colors Aspose.Cells | move chart to different worksheet .NET | Chart.NSeries.ChangeColors | ChartColorPaletteType Aspose.Cells | duplicate chart workbook | reposition chart cells Aspose.Cells | C# Aspose.Cells chart example
// Common Searches: how to copy a chart to another sheet with Aspose.Cells | Aspose.Cells change series colors after cloning | move cloned chart to new range C# | Aspose.Cells chart color palette index | duplicate worksheet with chart Aspose.Cells
// Developer Intent: The developer needs to duplicate an existing chart, apply a different color palette to its series, and place the modified chart on a separate worksheet without altering the original.
// Use Cases: Create a presentation‑ready summary sheet that contains a recolored copy of a source chart. | Generate regional reports by cloning a template chart and assigning each copy a distinct palette. | Automate report workflows where the original chart remains unchanged while a customized version is positioned on another sheet.
// AI Prompts: Write C# code using Aspose.Cells to clone a chart from one worksheet, set a custom RGB color for each series, and move it to a specified cell range on another worksheet. | Explain the purpose of Chart.NSeries.ChangeColors, how ChartColorPaletteType works, and how to select a specific palette index in Aspose.Cells. | Provide a step‑by‑step tutorial for copying a chart, recoloring its series, and repositioning it on a different sheet without affecting the source chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, duplicate the worksheet (which copies the chart), change the cloned chart's series colors using a built‑in palette, reposition the chart on the new sheet, and save the file.
class CloneChartExample
{
    static void Main()
    {
        // Create a new workbook and obtain the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate sample data for the chart
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["A2"].PutValue("Jan");
        sourceSheet.Cells["A3"].PutValue("Feb");
        sourceSheet.Cells["A4"].PutValue("Mar");

        sourceSheet.Cells["B1"].PutValue("Series1");
        sourceSheet.Cells["B2"].PutValue(10);
        sourceSheet.Cells["B3"].PutValue(20);
        sourceSheet.Cells["B4"].PutValue(30);

        sourceSheet.Cells["C1"].PutValue("Series2");
        sourceSheet.Cells["C2"].PutValue(15);
        sourceSheet.Cells["C3"].PutValue(25);
        sourceSheet.Cells["C4"].PutValue(35);

        // Add a column chart on the source sheet
        int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart sourceChart = sourceSheet.Charts[chartIdx];
        sourceChart.NSeries.Add("B1:C4", true);          // Set data range for series
        sourceChart.NSeries.CategoryData = "A2:A4";     // Set category (X‑axis) data

        // Clone the worksheet (which also clones the chart) to a new sheet
        int copiedIdx = workbook.Worksheets.AddCopy(0); // AddCopy by source index
        Worksheet clonedSheet = workbook.Worksheets[copiedIdx];
        clonedSheet.Name = "Cloned";

        // Retrieve the cloned chart (same index as the original chart)
        Chart clonedChart = clonedSheet.Charts[chartIdx];

        // Change the series colors using a monochromatic palette
        clonedChart.NSeries.ChangeColors((ChartColorPaletteType)0); // Palette index 0

        // Optionally move the cloned chart to a different location on the new sheet
        clonedChart.Move(10, 2, 25, 12); // topRow, leftColumn, bottomRow, rightColumn

        // Save the workbook with the cloned and recolored chart
        workbook.Save("ClonedChartDemo.xlsx");
    }
}
