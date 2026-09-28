// Title: Resizing data label shapes to a fixed width and height after enabling leader lines in an Aspose.Cells column chart (C#)
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells, activates data labels and leader lines, disables auto‑fit for each label, and sets the label shape to 80 pt width and 30 pt height. | Write a C# snippet using Aspose.Cells to set the leader line weight to 1 pt, color to dark gray, and apply a fixed‑size data label rectangle to every point in a chart series. | Provide a C# example that iterates over chart points, turns off DataLabels.IsResizeShapeToFitText, and assigns explicit Width and Height values to the data label shapes after enabling leader lines.
// Common Searches: Aspose.Cells set fixed width and height for chart data label shapes in C# | C# Aspose.Cells enable data label leader lines and prevent auto resizing | How to customize leader line color and weight in an Aspose.Cells column chart | Resize data label rectangles after turning on leader lines using Aspose.Cells | Aspose.Cells chart data labels fixed size without auto‑fit C#
// Tags: fixed-size data label shapes Aspose.Cells | activate data label leader lines Aspose.Cells | disable auto‑fit data labels Aspose.Cells | set data label dimensions column chart | customize leader line weight color Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to create a column chart with Aspose.Cells, enable data labels and leader lines, turn off automatic label sizing, and assign a fixed width of 80 points and height of 30 points to each data label shape before saving the workbook.
    public class ResizeDataLabelShapesWithLeaderLines
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
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable data labels and leader lines
                series.DataLabels.ShowValue = true;
                series.HasLeaderLines = true;                     // Enable leader lines
                series.LeaderLines.IsAuto = false;                // Disable automatic leader line layout
                // LeaderLines does not expose a LineStyle property; omit setting it
                series.LeaderLines.WeightPt = 1.0;
                series.LeaderLines.Color = Color.DarkGray;

                // Resize each data label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Prevent auto‑fit so we can set a fixed size
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Optionally set a custom shape type; default is rectangle, so this line is omitted
                    // point.DataLabels.ShapeType = MsoPresetShape.Rectangle;

                    // Set fixed dimensions (width and height are in points)
                    point.DataLabels.Width = 80;   // width in points
                    point.DataLabels.Height = 30;  // height in points
                }

                // Save the workbook
                string outputPath = "ResizeDataLabelShapesWithLeaderLines.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
