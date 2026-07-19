// Title: Aspose.Cells C# – Set Transparent Fill for All Chart Legend Entries
// Description: Creates a workbook, adds a column chart with two series, calculates the chart, then loops through the LegendEntryCollection to set each entry’s BackgroundMode to Transparent, producing a legend with no background color before saving the file.
// Keywords: Aspose.Cells | C# | chart legend transparent | LegendEntry BackgroundMode | transparent fill | remove legend shading | column chart | Excel workbook | Aspose.Cells API | chart formatting
// Common Searches: Aspose.Cells make legend background transparent C# | set chart legend entry fill to transparent Aspose.Cells | transparent legend entries Aspose.Cells example | how to remove legend shading in Aspose.Cells chart | Aspose.Cells LegendEntry BackgroundMode Transparent
// Developer Intent: I need to apply a transparent fill to every legend entry in an Aspose.Cells chart so the legend blends with the worksheet background.
// Use Cases: Design clean Excel reports where chart legends should not obscure background colors. | Prepare charts for PDF export where legend shading interferes with document design. | Create presentation‑ready workbooks with legends that match a custom slide background. | Standardize chart appearance across multiple workbooks by programmatically clearing legend fills.
// AI Prompts: Write C# code using Aspose.Cells to set BackgroundMode.Transparent for all LegendEntry objects in a chart. | Explain how to access and modify LegendEntryCollection for different chart types (column, line, pie) in Aspose.Cells. | Show how to verify legend transparency after saving the workbook, e.g., by opening the file or checking properties. | Provide a step‑by‑step guide to create a chart, calculate it, and apply transparent legend entries with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart with two series, calculates the chart, then loops through the LegendEntryCollection to set each entry’s BackgroundMode to Transparent, producing a legend with no background color before saving the file.
    public class TransparentLegendEntriesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");

                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(70);
                worksheet.Cells["B4"].PutValue(90);

                worksheet.Cells["C1"].PutValue("Series 2");
                worksheet.Cells["C2"].PutValue(60);
                worksheet.Cells["C3"].PutValue(80);
                worksheet.Cells["C4"].PutValue(100);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Add two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.CategoryData = "A2:A4";

                // Ensure the chart is calculated so legend entries are generated
                chart.Calculate();

                // Iterate over all legend entries and set their background to transparent
                LegendEntryCollection legendEntries = chart.Legend.LegendEntries;
                for (int i = 0; i < legendEntries.Count; i++)
                {
                    LegendEntry entry = legendEntries[i];
                    entry.BackgroundMode = BackgroundMode.Transparent; // Transparent fill
                }

                // Save the workbook
                string outputPath = "TransparentLegendEntriesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            TransparentLegendEntriesDemo.Run();
        }
    }
}
