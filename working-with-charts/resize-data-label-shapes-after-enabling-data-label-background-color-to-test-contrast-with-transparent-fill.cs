// Title: C# – Resize Chart Data Label Shapes with Transparent Background using Aspose.Cells
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, set a rectangular shape with a transparent fill, turn off automatic resizing, and apply a fixed 60 × 30 pixel size to both series‑level and point‑level data labels before saving the file.
// Keywords: Aspose.Cells C# chart data labels | custom data label size Aspose.Cells | transparent background data labels | disable data label auto‑fit | set data label width pixel | set data label height pixel | DataLabelShapeType.Rect | BackgroundMode.Transparent | ResizeDataLabelShapes example | GitHub Aspose.Cells chart sample
// Common Searches: how to set fixed size for chart data labels Aspose.Cells | Aspose.Cells data label transparent background | resize data label shape C# Aspose.Cells | disable auto‑fit for data labels .NET | apply custom pixel dimensions to chart data labels
// Developer Intent: The developer wants precise control over the dimensions and background of chart data label shapes instead of the default auto‑fit behavior.
// Use Cases: Create a column chart where every data label has a uniform 60 × 30 pixel rectangular shape with a transparent fill for visual contrast testing. | Iterate through each ChartPoint to enforce the same custom size and background on individual point labels. | Produce an Excel file with consistent data label dimensions regardless of label text length, improving layout consistency.
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart whose data labels use an 80 × 40 pixel oval shape and a semi‑transparent background. | Explain step‑by‑step how to disable automatic resizing of chart data label shapes and set custom pixel width and height for both series‑level and point‑level labels in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ResizeDataLabelShapesDemo
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels, set a rectangular shape with a transparent fill, turn off automatic resizing, and apply a fixed 60 × 30 pixel size to both series‑level and point‑level data labels before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(45);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and configure appearance
            DataLabels dataLabels = series.DataLabels;
            dataLabels.ShowValue = true;                         // Show the numeric value
            dataLabels.ShapeType = DataLabelShapeType.Rect;      // Use a rectangular shape
            dataLabels.BackgroundMode = BackgroundMode.Transparent; // Transparent fill for contrast testing

            // Disable automatic resizing so we can set a custom size
            dataLabels.IsResizeShapeToFitText = false;

            // Set a custom size that is smaller than the default auto‑fit size
            dataLabels.WidthPixel = 60;   // Width in pixels
            dataLabels.HeightPixel = 30;  // Height in pixels

            // Optionally, apply the same size to each individual point's label
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.IsResizeShapeToFitText = false;
                point.DataLabels.WidthPixel = 60;
                point.DataLabels.HeightPixel = 30;
                point.DataLabels.BackgroundMode = BackgroundMode.Transparent;
                point.DataLabels.ShapeType = DataLabelShapeType.Rect;
            }

            // Save the workbook
            workbook.Save("ResizeDataLabelShapes.xlsx");
        }
    }
}
