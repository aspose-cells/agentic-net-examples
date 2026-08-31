// Title: Make chart legend entry backgrounds transparent while preserving text color using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that iterates over chart.Legend.LegendEntries and sets each LegendEntry.BackgroundMode to Transparent while keeping the legend text fill enabled. | Show how to remove the background fill of legend entries in an Aspose.Cells chart without changing the existing text color.
// Common Searches: Aspose.Cells C# make chart legend background transparent | remove legend fill but keep text color Aspose.Cells chart | transparent legend entries Aspose.Cells example | how to set LegendEntry.BackgroundMode to Transparent in C# | chart legend customization without affecting text color Aspose.Cells
// Tags: legend entry background transparent Aspose.Cells | chart legend text fill preservation C# | Aspose.Cells set LegendEntry.BackgroundMode | remove legend fill without affecting text Aspose.Cells | transparent legend background chart Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendBackgroundRemoval
{
    // Creates a workbook, adds sample data, builds a column chart, makes the legend visible, iterates through each LegendEntry to set BackgroundMode to Transparent while ensuring the legend text retains its fill, and saves the file as ChartWithTransparentLegendBackground.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(70);
            sheet.Cells["C3"].PutValue(60);
            sheet.Cells["C4"].PutValue(90);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add series data
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is visible
            chart.ShowLegend = true;

            // Iterate through each legend entry and remove its background fill
            // while preserving the existing text color for contrast.
            foreach (LegendEntry entry in chart.Legend.LegendEntries)
            {
                // Make the background transparent (no fill)
                entry.BackgroundMode = BackgroundMode.Transparent;

                // Ensure the text itself still has fill (default behavior)
                entry.IsTextNoFill = false;
            }

            // Save the workbook
            workbook.Save("ChartWithTransparentLegendBackground.xlsx");
        }
    }
}
