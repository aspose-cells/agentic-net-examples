// Title: Aspose.Cells .NET – Rotate Chart Legend 90° and Preserve Transparent Entry Fill
// Description: This example creates a workbook, adds a column chart, shows the legend, sets its RotationAngle to 90 degrees, and iterates through each LegendEntry to keep the BackgroundMode transparent, then saves the file as LegendRotationTransparentFill.xlsx.
// Keywords: Aspose.Cells rotate legend | chart legend 90 degrees .NET | transparent legend entry background | Legend.RotationAngle Aspose | BackgroundMode.Transparent chart legend
// Common Searches: rotate chart legend 90 degrees Aspose.Cells | keep legend entry background transparent after rotation | Aspose.Cells legend rotation without changing fill | set legend rotation angle .NET Excel chart | transparent legend entries Aspose.Cells
// Developer Intent: Set a chart legend to 90° rotation while ensuring each legend entry stays transparent.
// Use Cases: Design Excel reports with vertical legends for space‑saving layouts. | Build dashboards where legend orientation must be upward without altering entry styling. | Generate automated spreadsheets that require a rotated legend while maintaining a clean, transparent background for each entry.
// AI Prompts: Generate C# code using Aspose.Cells to rotate a chart legend 90 degrees and keep legend entries transparent. | Show how to apply Legend.RotationAngle = 90 and set LegendEntry.BackgroundMode to Transparent in Aspose.Cells. | Explain the steps to rotate a chart legend without affecting its fill properties in a .NET Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendRotationDemo
{
    // This example creates a workbook, adds a column chart, shows the legend, sets its RotationAngle to 90 degrees, and iterates through each LegendEntry to keep the BackgroundMode transparent, then saves the file as LegendRotationTransparentFill.xlsx.
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is visible
            chart.ShowLegend = true;

            // Set legend rotation angle to 90 degrees (upward)
            Legend legend = chart.Legend;
            legend.RotationAngle = 90;

            // Preserve transparent fill for each legend entry
            foreach (LegendEntry entry in legend.LegendEntries)
            {
                // Keep background mode transparent (does not affect existing fill)
                entry.BackgroundMode = BackgroundMode.Transparent;
            }

            // Save the workbook
            workbook.Save("LegendRotationTransparentFill.xlsx");
        }
    }
}
