// Title: Aspose.Cells .NET – Rotate Chart Legend 90° and Keep Entry Background Transparent
// Description: Creates a workbook with a column chart, sets the legend rotation to 90 degrees, and makes each legend entry’s background transparent while preserving text fill, then saves the file.
// Keywords: Aspose.Cells rotate legend | chart legend 90 degrees .NET | transparent legend entry background | Legend.RotationAngle Aspose | BackgroundMode.Transparent chart | C# Excel chart formatting
// Common Searches: rotate chart legend 90° Aspose.Cells | keep legend entry background transparent after rotation | Aspose.Cells legend formatting C# | set legend rotation angle without changing fill | transparent legend entries Excel chart
// Developer Intent: Set a chart legend to vertical orientation while retaining transparent entry backgrounds.
// Use Cases: Display a vertical legend in narrow charts without visual clutter. | Generate Excel reports where legends must be rotated for layout constraints. | Apply uniform transparent styling to legend entries across multiple charts programmatically.
// AI Prompts: Write C# code using Aspose.Cells to rotate a chart legend 90 degrees and ensure each legend entry has a transparent background. | Explain how to change Legend.RotationAngle and set LegendEntry.BackgroundMode to Transparent without affecting text appearance. | Show a method to apply transparent legend entry backgrounds to several charts in a workbook after rotating their legends.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendRotationDemo
{
    // Creates a workbook with a column chart, sets the legend rotation to 90 degrees, and makes each legend entry’s background transparent while preserving text fill, then saves the file.
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

            // Ensure each legend entry keeps a transparent background fill
            foreach (LegendEntry entry in chart.Legend.LegendEntries)
            {
                entry.BackgroundMode = BackgroundMode.Transparent; // transparent fill
                entry.IsTextNoFill = false;                       // text fill remains enabled
            }

            // Save the workbook
            workbook.Save("LegendRotationTransparentFill.xlsx");
        }
    }
}
