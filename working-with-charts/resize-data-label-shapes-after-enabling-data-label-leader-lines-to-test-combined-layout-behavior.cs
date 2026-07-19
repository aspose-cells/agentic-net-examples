// Title: Resize Data Label Shapes with Leader Lines in an Aspose.Cells Column Chart (C#)
// Description: Demonstrates how to create a column chart in Aspose.Cells, enable leader lines for a series, turn on data labels, disable automatic shape resizing, set explicit pixel width and height, apply a RightArrowCallout shape, and save the workbook.
// Keywords: Aspose.Cells data label resize | C# chart leader lines | fixed size data labels Aspose.Cells | custom data label shape callout | column chart label dimensions | disable auto resize Aspose.Cells | set data label width pixel | set data label height pixel | Aspose.Cells chart customization
// Common Searches: Aspose.Cells set data label width and height | how to disable automatic data label resizing in Aspose.Cells | leader lines with custom callout shape Aspose.Cells C# | fixed-size data labels for Excel chart using Aspose.Cells | resize data label shapes after enabling leader lines
// Developer Intent: The developer wants to prevent automatic resizing of chart data label shapes, assign specific pixel dimensions, and use a callout shape while leader lines are active.
// Use Cases: Design a dashboard Excel file where column chart labels keep a uniform size regardless of value length. | Create presentation‑ready charts with leader lines and callout labels that do not shift layout when data changes. | Generate automated reports where label dimensions are controlled to match corporate branding guidelines.
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart, enables leader lines, and sets data label width to 80 px and height to 40 px using a RightArrowCallout shape. | Explain step‑by‑step how to turn off automatic data label resizing and apply a custom callout shape in Aspose.Cells charts. | Provide a snippet to customize leader line style, weight, and color for a series in an Aspose.Cells chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a column chart in Aspose.Cells, enable leader lines for a series, turn on data labels, disable automatic shape resizing, set explicit pixel width and height, apply a RightArrowCallout shape, and save the workbook.
    class ResizeDataLabelShapesDemo
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

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series and enable leader lines
                Series series = chart.NSeries[0];
                series.HasLeaderLines = true;
                // Optional: customize the appearance of the leader lines
                series.LeaderLines.IsAuto = false;
                series.LeaderLines.Style = LineType.Solid;
                series.LeaderLines.WeightPt = 1.0;
                series.LeaderLines.Color = Color.DarkGray;

                // Enable data labels for the series
                DataLabels dataLabels = series.DataLabels;
                dataLabels.ShowValue = true;
                dataLabels.Position = LabelPositionType.OutsideEnd;

                // Disable automatic resizing so we can set explicit dimensions
                dataLabels.IsResizeShapeToFitText = false;
                // Set custom size (smaller than the text would normally require)
                dataLabels.WidthPixel = 60;   // width in pixels
                dataLabels.HeightPixel = 30;  // height in pixels

                // Use a callout shape type to better illustrate leader line behavior
                dataLabels.ShapeType = DataLabelShapeType.RightArrowCallout;

                // Save the workbook
                string outputPath = "ResizeDataLabelShapesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ResizeDataLabelShapesDemo.Run();
        }
    }
}
