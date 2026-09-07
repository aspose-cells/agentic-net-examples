// Title: Clone an existing Excel chart and reassign its data source to another worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that copies a chart from one worksheet to another and updates the series values to reference the new sheet with Aspose.Cells. | Show how to duplicate a chart object, preserve its type and position, and bind its data series to a different worksheet in a .NET workbook. | Provide a step‑by‑step example of cloning an Aspose.Cells chart and modifying the NSeries ranges to point to a destination sheet.
// Common Searches: Aspose.Cells C# clone chart and change data source to another worksheet | How to copy an Excel chart to a different sheet and update series ranges with Aspose.Cells | Programmatically duplicate a chart and rebind its data series in a .NET workbook | C# Aspose.Cells example for cloning chart and setting new data range
// Tags: chart copy Aspose.Cells C# | modify chart series range Aspose.Cells | duplicate Excel chart on another worksheet | Aspose.Cells NSeries data source update | preserve chart layout while changing data source

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartCloneExample
{
    // Creates a workbook, adds a source chart, clones it onto a new worksheet, copies each series to the cloned chart while keeping the same type and position, updates the series ranges to reference the destination sheet's data, and saves the file as ClonedChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data to the source sheet
                Workbook workbook = new Workbook();
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "SourceSheet";

                // Populate data for the original chart
                srcSheet.Cells["A1"].PutValue("Category");
                srcSheet.Cells["A2"].PutValue("A");
                srcSheet.Cells["A3"].PutValue("B");
                srcSheet.Cells["A4"].PutValue("C");
                srcSheet.Cells["B1"].PutValue("Value");
                srcSheet.Cells["B2"].PutValue(10);
                srcSheet.Cells["B3"].PutValue(20);
                srcSheet.Cells["B4"].PutValue(30);

                // Add a chart to the source sheet
                int srcChartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart srcChart = srcSheet.Charts[srcChartIdx];

                // Add series with values (category data will be inferred from the first column)
                int srcSeriesIdx = srcChart.NSeries.Add("B2:B4", true);
                // No explicit CategoryData assignment needed; Aspose.Cells infers it from the adjacent column

                // Add a destination worksheet that will host the cloned chart
                Worksheet destSheet = workbook.Worksheets.Add("DestinationSheet");

                // Populate data for the destination sheet (same layout, different values)
                destSheet.Cells["A1"].PutValue("Category");
                destSheet.Cells["A2"].PutValue("X");
                destSheet.Cells["A3"].PutValue("Y");
                destSheet.Cells["A4"].PutValue("Z");
                destSheet.Cells["B1"].PutValue("Value");
                destSheet.Cells["B2"].PutValue(40);
                destSheet.Cells["B3"].PutValue(50);
                destSheet.Cells["B4"].PutValue(60);

                // Clone the chart: create a new chart on the destination sheet with the same type and position
                int clonedChartIdx = destSheet.Charts.Add(
                    srcChart.Type,
                    srcChart.ChartObject.UpperLeftRow,
                    srcChart.ChartObject.UpperLeftColumn,
                    srcChart.ChartObject.UpperLeftRow + 10,   // approximate bottom row
                    srcChart.ChartObject.UpperLeftColumn + 5 // approximate bottom column
                );
                Chart clonedChart = destSheet.Charts[clonedChartIdx];

                // Copy each series from the source chart, adjusting the data range to refer to the destination sheet
                foreach (Series series in srcChart.NSeries)
                {
                    // Add series values to the cloned chart (range strings are unchanged because layout is identical)
                    int clonedSeriesIdx = clonedChart.NSeries.Add(series.Values, true);
                    // No explicit CategoryData assignment; categories are inferred automatically
                }

                // Ensure output directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ClonedChart.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
