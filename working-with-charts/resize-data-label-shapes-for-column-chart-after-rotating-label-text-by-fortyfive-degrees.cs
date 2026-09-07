// Title: Set fixed width and height for 45-degree rotated data label shapes in a column chart with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that rotates the data labels of a column chart by 45 degrees and assigns a fixed width of 80 px and height of 40 px to each label shape. | Show how to turn off automatic resizing of data label shapes and apply the same custom dimensions to every point in a column chart with Aspose.Cells for .NET. | Provide a complete example that creates a workbook, adds a column chart, configures rotated data labels with custom shape size, and saves the file.
// Common Searches: Aspose.Cells how to set fixed width and height for rotated data labels in a column chart | C# disable automatic data label shape resizing in Aspose.Cells chart | rotate data label text 45 degrees and adjust label box size Aspose.Cells .NET | apply same data label dimensions to all points in a column chart using Aspose.Cells
// Tags: rotate column chart data labels Aspose.Cells | custom data label shape size Aspose.Cells | disable data label auto resize Aspose.Cells | set data label width height .NET | column chart label dimensions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds a column chart with sample data, enables data labels, rotates them 45°, disables automatic shape resizing, sets a fixed width of 80 px and height of 40 px for the series and each point's data labels, and saves the workbook as ResizeRotatedDataLabels.xlsx.
    public class ResizeRotatedDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                DataLabels seriesLabels = chart.NSeries[0].DataLabels;
                seriesLabels.ShowValue = true;

                // Rotate the data label text by 45 degrees
                seriesLabels.RotationAngle = 45;

                // Disable automatic shape resizing so we can set custom dimensions
                seriesLabels.IsResizeShapeToFitText = false;

                // Set a custom width and height that fits the rotated text (pixels)
                seriesLabels.Width = 80;
                seriesLabels.Height = 40;

                // Apply the same settings to each individual point (optional, ensures consistency)
                foreach (ChartPoint point in chart.NSeries[0].Points)
                {
                    point.DataLabels.RotationAngle = 45;
                    point.DataLabels.IsResizeShapeToFitText = false;
                    point.DataLabels.Width = 80;
                    point.DataLabels.Height = 40;
                }

                // Save the workbook
                workbook.Save("ResizeRotatedDataLabels.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            ResizeRotatedDataLabels.Run();
        }
    }
}
