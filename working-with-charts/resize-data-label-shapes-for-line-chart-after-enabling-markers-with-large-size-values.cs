// Title: Resize data label shapes for a line chart with large markers using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart, sets a circular marker size to 20 points, disables auto‑fit for each point's data label, and assigns a 60 pt width and 20 pt height to the label shapes. | Demonstrate how to loop through ChartPoint objects in Aspose.Cells to customize data label dimensions after enabling large markers on a line series.
// Common Searches: Aspose.Cells C# set fixed width and height for line chart data labels | disable auto resize of data label shapes in Aspose.Cells line chart | increase marker size and customize data label shape size Aspose.Cells | prevent data label auto‑fit in Aspose.Cells chart series | Aspose.Cells example resize label shape points XLSX
// Tags: Aspose.Cells line chart data label size | custom data label dimensions C# | disable data label auto‑fit Aspose.Cells | large marker size line chart Aspose.Cells | set chart point label shape size | Aspose.Cells resize label shape XLSX

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds a line chart with sample data, enables a large circular marker, turns on data labels, disables auto‑fit for each point's label, sets each label shape to 60 pt width and 20 pt height, and saves the file as an XLSX workbook.
    public class ResizeDataLabelShapesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
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

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first (and only) series
                Series series = chart.NSeries[0];

                // Enable markers and set a large marker size
                series.Marker.MarkerStyle = ChartMarkerType.Circle;
                series.Marker.MarkerSize = 20; // large size in points

                // Enable data labels for the series
                series.DataLabels.ShowValue = true;

                // For each point, disable auto‑fit and set a custom width/height for the label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Prevent the shape from auto‑resizing to fit the text
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set custom dimensions (units are points)
                    point.DataLabels.Width = 60;   // make the label wider
                    point.DataLabels.Height = 20;  // make the label taller
                }

                // Define output file path
                string outputPath = "ResizeDataLabelShapesDemo.xlsx";

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

    public class Program
    {
        public static void Main(string[] args)
        {
            // Entry point for the application
            ResizeDataLabelShapesDemo.Run();
        }
    }
}
