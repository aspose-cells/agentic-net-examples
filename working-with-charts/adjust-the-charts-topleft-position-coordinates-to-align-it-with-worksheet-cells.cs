// Title: Align an Aspose.Cells chart to specific worksheet cells by setting its top‑left and bottom‑right coordinates in C#
// AI Prompts: Generate C# code that moves an Aspose.Cells chart so its top‑left corner is anchored to cell C2 and its bottom‑right corner to cell H12. | Show how to set the Placement property to MoveAndSize after positioning a chart with the Chart.Move method in Aspose.Cells. | Provide a step‑by‑step example of attaching a column chart to a range of rows and columns using Aspose.Cells. | Explain how to ensure a chart resizes with the worksheet after aligning it to a cell range in Aspose.Cells C#.
// Common Searches: Aspose.Cells C# align chart to cell C2 and H12 | how to use Chart.Move to position a chart in Aspose.Cells | set chart placement MoveAndSize Aspose.Cells example | anchor Aspose.Cells chart to a specific cell range | adjust chart top left and bottom right coordinates in Aspose.Cells
// Tags: chart anchoring to worksheet cells Aspose.Cells | move and size chart behavior Aspose.Cells | column chart cell range alignment Aspose.Cells | set chart bounds by rows and columns Aspose.Cells | Aspose.Cells chart positioning example C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, inserts a column chart, anchors the chart's top‑left corner to cell C2 and bottom‑right corner to cell H12 using Chart.Move, sets the chart's Placement to MoveAndSize so it moves and resizes with the cells, and saves the file as AlignedChart.xlsx.
    public class AlignChartToCells
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
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["A4"].PutValue("Cherry");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(40);
                worksheet.Cells["B3"].PutValue(55);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart. Initial position is arbitrary (rows 5‑15, columns 0‑5)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Align the chart's top‑left corner to cell C2 (row index 1, column index 2)
                // and its bottom‑right corner to cell H12 (row index 11, column index 7)
                chart.Move(topRow: 1, leftColumn: 2, bottomRow: 11, rightColumn: 7);

                // Make the chart move and size with the cells it is attached to
                chart.Placement = PlacementType.MoveAndSize;

                // Save the workbook
                string outputPath = "AlignedChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AlignChartToCells.Run();
        }
    }
}
