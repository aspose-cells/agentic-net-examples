// Title: Rotate chart legend 90° and keep legend entries transparent using Aspose.Cells for .NET
// Description: C# example that creates a column chart, sets Legend.RotationAngle to 90 degrees, and makes each LegendEntry background transparent without altering text fill, then saves the workbook as an Excel file.
// Keywords: Aspose.Cells chart legend rotation | C# Legend.RotationAngle 90 | transparent legend entries Aspose.Cells | BackgroundMode.Transparent chart legend | .NET Excel chart customization | Aspose.Cells LegendEntry background | rotate legend vertical Aspose.Cells
// Common Searches: Aspose.Cells rotate chart legend 90 degrees | make legend entries transparent Aspose.Cells .NET | set Legend.RotationAngle in C# Aspose.Cells | how to keep legend background transparent in Excel chart | Aspose.Cells legend entry background mode transparent
// Developer Intent: Apply a 90‑degree rotation to a chart legend while preserving a transparent background for each legend entry.
// Use Cases: Design compact dashboards where a vertical legend saves horizontal space. | Generate clean‑look Excel reports with legend boxes that have no fill, ideal for printing or PDF export. | Automate consistent legend styling across multiple charts in a reporting pipeline.
// AI Prompts: Show C# code to set Legend.RotationAngle = 90 and make LegendEntry.BackgroundMode Transparent in Aspose.Cells. | Explain how to iterate over LegendEntries to apply a transparent background without changing text fill. | Provide a step‑by‑step guide for rotating a chart legend vertically while keeping legend entry fill transparent using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendRotationExample
{
    // C# example that creates a column chart, sets Legend.RotationAngle to 90 degrees, and makes each LegendEntry background transparent without altering text fill, then saves the workbook as an Excel file.
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

            // Rotate the legend text by 90 degrees (upward)
            chart.Legend.RotationAngle = 90;

            // Ensure each legend entry retains a transparent background (no fill)
            foreach (LegendEntry entry in chart.Legend.LegendEntries)
            {
                entry.BackgroundMode = BackgroundMode.Transparent;
                // Preserve existing IsTextNoFill setting (do not modify)
            }

            // Save the workbook
            workbook.Save("LegendRotationTransparentFill.xlsx");
        }
    }
}
