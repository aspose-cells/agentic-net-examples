// Title: Aspose.Cells for .NET – Set Chart Legend to Bottom‑Right Corner (C#)
// Description: This example creates a workbook, adds sample data and a column chart, then loops through all charts on the worksheet and assigns `Legend.Position = LegendPositionType.Corner`. The legend is placed in the bottom‑right corner of the plot area, keeping it out of the data series before the file is saved.
// Keywords: Aspose.Cells chart legend position | C# set legend bottom right | LegendPositionType.Corner | Aspose.Cells .NET chart formatting | prevent legend overlap
// Common Searches: Aspose.Cells set legend to corner C# | how to move chart legend bottom right Aspose.Cells | C# chart legend placement Aspose.Cells example | Aspose.Cells legend position property
// Developer Intent: Move every chart legend to the bottom‑right corner so it does not cover the plotted data.
// Use Cases: Standardize legend placement across multiple charts in a financial dashboard. | Generate automated reports where legends must stay outside the data area for clarity. | Update existing workbooks to reposition legends without recreating charts.
// AI Prompts: Show C# code using Aspose.Cells to set all chart legends to the bottom‑right corner. | Give an Aspose.Cells example that iterates through a worksheet's Charts collection and applies LegendPositionType.Corner. | Explain how to change the legend position of a saved workbook without adding new charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart, ChartType, LegendPositionType

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds sample data and a column chart, then loops through all charts on the worksheet and assigns `Legend.Position = LegendPositionType.Corner`. The legend is placed in the bottom‑right corner of the plot area, keeping it out of the data series before the file is saved.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for demonstration
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];
                chart.SetChartDataRange("A1:B4", true);

                // Set legend position for all charts in the worksheet to the bottom‑right corner
                foreach (Chart c in sheet.Charts)
                {
                    c.Legend.Position = LegendPositionType.Corner; // bottom‑right corner of plot area
                    // Aspose.Cells does not expose an IsOverlaid property; the default behavior does not overlay the chart data.
                }

                // Save the workbook
                string outputPath = "ChartsWithCornerLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
