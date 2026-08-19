// Title: Make Chart Legend Fully Transparent in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, and render the legend completely invisible by setting `chart.Legend.BackgroundMode` to `BackgroundMode.Transparent` and configuring each `LegendEntry` with `IsTextNoFill = true` and a transparent background. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells | C# chart legend transparent | BackgroundMode.Transparent | legend entry no fill | transparent Excel legend | Aspose.Cells chart styling | clear legend Aspose.Cells
// Common Searches: Aspose.Cells make chart legend transparent C# | remove legend background in Excel using Aspose.Cells | set legend entry no fill Aspose.Cells .NET | transparent legend Aspose.Cells example | how to hide chart legend in Aspose.Cells
// Developer Intent: Render a chart legend and its entries with no visible background or fill in an Aspose.Cells workbook.
// Use Cases: Design minimalist Excel reports where the legend should not distract from the data. | Overlay charts on images or colored cells without the legend obscuring the view. | Create dashboards that require legends to blend seamlessly with the worksheet background.
// AI Prompts: Provide C# code using Aspose.Cells to set a chart legend and all legend entries to transparent with no text fill. | Explain step‑by‑step how to apply BackgroundMode.Transparent to a chart legend and its entries in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add sample data, insert a column chart, and render the legend completely invisible by setting `chart.Legend.BackgroundMode` to `BackgroundMode.Transparent` and configuring each `LegendEntry` with `IsTextNoFill = true` and a transparent background. The workbook is then saved as an Excel file.
class TransparentLegendDemo
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Make the legend background fully transparent
        chart.Legend.BackgroundMode = BackgroundMode.Transparent;

        // Ensure each legend entry has no fill (text fill disabled) and also transparent background
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            entry.IsTextNoFill = true;                     // No fill for the text
            entry.BackgroundMode = BackgroundMode.Transparent; // Transparent entry background
        }

        // Save the workbook
        workbook.Save("TransparentLegend.xlsx");
    }
}
