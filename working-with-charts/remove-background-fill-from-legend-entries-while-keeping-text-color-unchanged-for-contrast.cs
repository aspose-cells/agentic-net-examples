// Title: C# – Remove Chart Legend Background Fill (keep text color) with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a column chart, enable its legend, and set each LegendEntry's BackgroundMode to Transparent so the legend background disappears while the label colors stay visible. The workbook is saved as LegendWithoutBackgroundFill.xlsx.
// Keywords: Aspose.Cells legend transparent | C# chart legend background | remove legend fill Aspose.Cells | Aspose.Cells BackgroundMode Transparent | Excel chart legend styling .NET
// Common Searches: Aspose.Cells make legend background transparent | C# remove fill from chart legend entries | how to keep legend text color while clearing background in Aspose.Cells | set legend entry background mode to transparent .NET
// Developer Intent: Strip the background color from all legend entries of an Excel chart while preserving the original text color for readability.
// Use Cases: Design clean Excel reports where legends blend with sheet backgrounds. | Generate charts for presentations that require no legend fill to avoid visual clutter. | Apply a uniform transparent legend style across multiple charts in a single workbook.
// AI Prompts: Write C# code using Aspose.Cells to set every LegendEntry.BackgroundMode to Transparent for any chart type. | Show how to loop through a LegendEntryCollection and change only the background fill, leaving font settings untouched. | Explain step‑by‑step how to customize legend appearance—removing fill and retaining font color—in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, enable its legend, and set each LegendEntry's BackgroundMode to Transparent so the legend background disappears while the label colors stay visible. The workbook is saved as LegendWithoutBackgroundFill.xlsx.
    public class RemoveLegendBackgroundFill
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(50);
                sheet.Cells["B4"].PutValue(70);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Ensure the legend is visible
                chart.ShowLegend = true;

                // Iterate over all legend entries and remove their background fill
                LegendEntryCollection entries = chart.Legend.LegendEntries;
                foreach (LegendEntry entry in entries)
                {
                    // Set background mode to Transparent to remove fill
                    entry.BackgroundMode = BackgroundMode.Transparent;
                }

                // Save the workbook
                workbook.Save("LegendWithoutBackgroundFill.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveLegendBackgroundFill.Run();
        }
    }
}
