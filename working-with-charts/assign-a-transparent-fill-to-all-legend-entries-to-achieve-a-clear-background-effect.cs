// Title: Aspose.Cells C# – Set Transparent Fill for All Chart Legend Entries
// Description: The sample builds an Excel workbook, inserts a column chart with two data series, forces chart calculation to create legend entries, and then assigns BackgroundMode.Transparent to each LegendEntry, resulting in a legend with no background before the file is saved.
// Keywords: Aspose.Cells | C# | chart legend transparent | BackgroundMode.Transparent | transparent legend fill | Excel chart formatting | remove legend background | Aspose.Cells LegendEntry | transparent chart legend | Aspose.Cells chart styling
// Common Searches: Aspose.Cells set legend entry transparent | C# transparent chart legend Aspose.Cells | how to remove legend background color in Aspose.Cells | BackgroundMode.Transparent chart legend | make chart legend background invisible Aspose.Cells
// Developer Intent: Apply a transparent fill to every legend entry of an Aspose.Cells chart so the legend appears without a background color.
// Use Cases: Design a dashboard workbook where the legend must blend seamlessly with colored cells. | Create presentation‑ready Excel reports that require a clean, unobtrusive legend overlay. | Generate multi‑series charts for web‑based analytics where the legend should not obscure underlying data visuals.
// AI Prompts: Provide C# code that iterates over chart.Legend.LegendEntries and sets each entry's BackgroundMode to Transparent using Aspose.Cells. | Explain why invoking chart.Calculate() is required before accessing LegendEntries when applying a transparent fill. | Show examples of applying a transparent legend background to pie, line, and bar charts with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The sample builds an Excel workbook, inserts a column chart with two data series, forces chart calculation to create legend entries, and then assigns BackgroundMode.Transparent to each LegendEntry, resulting in a legend with no background before the file is saved.
    public class TransparentLegendEntriesDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";     // Categories

            // Ensure the chart is calculated so legend entries are generated
            chart.Calculate();

            // Set each legend entry background to transparent
            LegendEntryCollection legendEntries = chart.Legend.LegendEntries;
            for (int i = 0; i < legendEntries.Count; i++)
            {
                LegendEntry entry = legendEntries[i];
                entry.BackgroundMode = BackgroundMode.Transparent; // Transparent fill
            }

            // Save the workbook
            workbook.Save("TransparentLegendEntriesDemo.xlsx");
        }
    }
}
