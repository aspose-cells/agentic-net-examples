using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ResizeDataLabelShapesForLineChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the line chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a line chart
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first (and only) series
                Series series = chart.NSeries[0];

                // Enable markers with a large size
                series.Marker.MarkerStyle = ChartMarkerType.Circle;
                series.Marker.MarkerSize = 30; // large marker size in points
                series.Marker.ForegroundColor = Color.Red;
                series.Marker.BackgroundColor = Color.Yellow;

                // Enable data labels for the series
                series.DataLabels.ShowValue = true;
                // Optional: set shape type if needed (removed due to API compatibility)

                // Resize each data label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Disable auto‑fit so we can set custom dimensions
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set custom width and height (units are points)
                    point.DataLabels.Width = 40;   // narrower than default
                    point.DataLabels.Height = 15;  // shorter than default
                }

                // Define output file path
                string outputPath = "ResizeDataLabelShapesForLineChart.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesForLineChart.Run();
        }
    }
}