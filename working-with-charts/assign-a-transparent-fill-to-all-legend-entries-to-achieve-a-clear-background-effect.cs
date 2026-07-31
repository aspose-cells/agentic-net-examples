// Title: Aspose.Cells C# – Set Chart Legend Background to Transparent
// Description: Creates a workbook, adds sample data and a column chart, then loops through the chart’s LegendEntries collection and assigns BackgroundMode.Transparent to each entry, resulting in a legend with no fill before the file is saved.
// Keywords: Aspose.Cells transparent legend | C# chart legend background mode | BackgroundMode.Transparent | remove legend fill Excel | Aspose.Cells legend styling | transparent chart legend Aspose | Excel legend background transparent C#
// Common Searches: Aspose.Cells make legend transparent | C# set legend background mode transparent | how to remove legend fill in Aspose.Cells chart | transparent legend entries Aspose.Cells example | chart legend background transparent C#
// Developer Intent: Apply a transparent fill to every legend entry in an Aspose.Cells chart using C#.
// Use Cases: Produce Excel dashboards where the legend must blend with custom backgrounds. | Generate presentation‑ready reports without legend shading obscuring data points. | Standardize styling across multiple charts in a workbook that require clear legends.
// AI Prompts: Write C# code with Aspose.Cells that sets BackgroundMode.Transparent for all LegendEntry objects in any chart type. | Explain how BackgroundMode.Transparent affects legend appearance and why text remains visible. | Show how to apply a transparent legend background to several charts in the same workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsLegendTransparentFill
{
    // Creates a workbook, adds sample data and a column chart, then loops through the chart’s LegendEntries collection and assigns BackgroundMode.Transparent to each entry, resulting in a legend with no fill before the file is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(70);
            sheet.Cells["B4"].PutValue(50);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the chart is calculated so legend entries are generated
            chart.Calculate();

            // Iterate through all legend entries and set a transparent background
            LegendEntryCollection legendEntries = chart.Legend.LegendEntries;
            for (int i = 0; i < legendEntries.Count; i++)
            {
                LegendEntry entry = legendEntries[i];
                entry.BackgroundMode = BackgroundMode.Transparent; // Transparent fill
                // Optional: ensure text itself has no fill (keeps text visible)
                entry.IsTextNoFill = false;
            }

            // Save the workbook
            workbook.Save("LegendTransparentFill.xlsx");
        }
    }
}
