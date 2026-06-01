using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class ResizeDataLabelShapesWithLeaderLines
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart data range
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable leader lines for the series
                series.HasLeaderLines = true;
                series.LeaderLines.IsAuto = false;
                series.LeaderLines.Style = LineType.Solid;
                series.LeaderLines.WeightPt = 1.0;
                series.LeaderLines.Color = Color.DarkGray;

                // Enable data labels and set basic properties
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.OutsideEnd;
                series.DataLabels.ShapeType = DataLabelShapeType.Rect;

                // Disable automatic shape resizing so we can set custom dimensions
                series.DataLabels.IsResizeShapeToFitText = false;

                // Set a custom size that is smaller than the default auto‑fit size
                series.DataLabels.Width = 60;   // pixels
                series.DataLabels.Height = 30;  // pixels

                // Apply the same settings to each individual point (optional, ensures per‑point consistency)
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.IsResizeShapeToFitText = false;
                    point.DataLabels.Width = 60;
                    point.DataLabels.Height = 30;
                    point.DataLabels.ShapeType = DataLabelShapeType.Rect;
                }

                // Define output file path
                string outputPath = "ResizeDataLabelShapesWithLeaderLines.xlsx";

                // Save the workbook
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
            ResizeDataLabelShapesWithLeaderLines.Run();
        }
    }
}