// Title: Make every chart legend entry transparent in an Aspose.Cells .NET (C#) workbook
// AI Prompts: Write C# code that loops through a chart's LegendEntryCollection in Aspose.Cells and sets each entry's BackgroundMode to Transparent. | Show how to apply a clear background to all legend entries of a column chart using Aspose.Cells for .NET. | Provide a complete example that creates a workbook, adds a chart, and makes the legend entries transparent with the Aspose.Cells C# API. | Demonstrate setting LegendEntry.BackgroundMode = BackgroundMode.Transparent for multiple legend entries in Aspose.Cells.
// Common Searches: aspose.cells make chart legend background transparent c# | c# set legend entry background mode transparent aspose cells | how to remove legend background color in aspose.cells chart | transparent legend entries example aspose.cells .net | aspose.cells legend transparency column chart c#
// Tags: chart legend background clear Aspose.Cells | set LegendEntry BackgroundMode Transparent C# | Aspose.Cells column chart legend styling | iterate LegendEntryCollection Aspose.Cells | legend entry fill transparent Aspose.Cells | clear legend entry fill Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds a column chart with sample data, iterates over the chart's LegendEntryCollection, sets each LegendEntry's BackgroundMode to Transparent, and saves the file as TransparentLegendEntriesDemo.xlsx.
    public class TransparentLegendEntriesDemo
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully: TransparentLegendEntriesDemo.xlsx");
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(80);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set each legend entry background to transparent
            LegendEntryCollection legendEntries = chart.Legend.LegendEntries;
            for (int i = 0; i < legendEntries.Count; i++)
            {
                LegendEntry entry = legendEntries[i];
                entry.BackgroundMode = BackgroundMode.Transparent;
            }

            // Save the workbook
            workbook.Save("TransparentLegendEntriesDemo.xlsx");
        }
    }
}
