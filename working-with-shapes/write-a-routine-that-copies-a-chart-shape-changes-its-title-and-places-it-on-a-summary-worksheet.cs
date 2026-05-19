using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    public class ChartCopyRoutine
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["A2"].PutValue("Apple");
                sourceSheet.Cells["A3"].PutValue("Banana");
                sourceSheet.Cells["A4"].PutValue("Cherry");
                sourceSheet.Cells["B1"].PutValue("Value");
                sourceSheet.Cells["B2"].PutValue(30);
                sourceSheet.Cells["B3"].PutValue(45);
                sourceSheet.Cells["B4"].PutValue(25);

                // Add a column chart to the source worksheet
                int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sourceSheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Get the ChartShape (the visual shape of the chart)
                ChartShape sourceChartShape = chart.ChartObject;

                // Ensure a worksheet named "Summary" exists
                Worksheet summarySheet = workbook.Worksheets["Summary"];
                if (summarySheet == null)
                {
                    summarySheet = workbook.Worksheets.Add("Summary");
                }

                // Copy the chart shape to the summary worksheet at a new location
                // Parameters: source shape, top row, vertical offset (pixels), left column, horizontal offset (pixels)
                Shape copiedShape = summarySheet.Shapes.AddCopy(sourceChartShape, 2, 0, 2, 0);

                // Change the title (caption) of the copied shape
                copiedShape.Title = "Summary Chart";

                // Optionally adjust size or position of the copied shape
                copiedShape.Width = sourceChartShape.Width;
                copiedShape.Height = sourceChartShape.Height;
                copiedShape.Top = 50;   // pixels from the top of the row
                copiedShape.Left = 100; // pixels from the left of the column

                // Save the workbook
                string outputPath = "ChartCopySummary.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ChartCopyRoutine.Run();
        }
    }
}