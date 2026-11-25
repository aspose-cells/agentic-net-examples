// Title: Resize bubble chart data label shapes after increasing bubble scale with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# resize bubble chart data label shape after setting BubbleScale | C# set custom width and height for bubble chart point data labels using Aspose.Cells | disable auto‑fit of data label shapes in an Aspose.Cells bubble chart | change background color of bubble chart data labels with Aspose.Cells .NET | increase bubble size scaling and adjust label dimensions in Aspose.Cells
// Tags: bubble chart data label custom size Aspose.Cells | set BubbleScale property Aspose.Cells | disable IsResizeShapeToFitText Aspose.Cells | custom data label background Aspose.Cells | C# chart point label dimensions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace BubbleChartDataLabelResize
{
    // The example creates a workbook, fills cells with X, Y, and size data, adds a bubble chart, sets the series BubbleScale to 180%, enables data labels, disables automatic shape resizing, assigns a fixed width of 70 points and height of 25 points to each point's label, applies a light‑yellow opaque background with a dark‑gray border, and saves the file as BubbleChart_DataLabelsResized.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Populate sample data for a bubble chart
            // ------------------------------------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Size");

            // Sample data (5 points)
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(10);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C4"].PutValue(15);

            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(25);
            sheet.Cells["C5"].PutValue(8);

            sheet.Cells["A6"].PutValue(5);
            sheet.Cells["B6"].PutValue(15);
            sheet.Cells["C6"].PutValue(12);

            // ------------------------------------------------------------
            // 2. Add a bubble chart
            // ------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 8, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series and bind X, Y and bubble size ranges
            int seriesIndex = chart.NSeries.Add("B2:B6", true); // Y values
            Series series = chart.NSeries[seriesIndex];
            series.XValues = "A2:A6";
            series.BubbleSizes = "C2:C6";

            // Increase bubble sizes globally using BubbleScale (e.g., 180%)
            series.BubbleScale = 180;

            // Optional: give the chart a title
            chart.Title.Text = "Bubble Chart – Data Labels Resized After Scaling";

            // ------------------------------------------------------------
            // 3. Resize data label shapes for each point
            // ------------------------------------------------------------
            // Ensure data labels are visible
            series.DataLabels.ShowValue = true;
            series.DataLabels.IsAutoText = true;

            // Iterate over each point in the series
            foreach (ChartPoint point in series.Points)
            {
                // Access the DataLabels object for the point
                DataLabels dl = point.DataLabels;

                // Disable automatic shape resizing to fit the text
                dl.IsResizeShapeToFitText = false;

                // Set a custom size for the label shape (width & height in points)
                // These values are deliberately smaller than the default auto‑fit size
                dl.Width = 70;   // width in points
                dl.Height = 25;  // height in points

                // Optionally change the background to make the label visible
                dl.BackgroundMode = BackgroundMode.Opaque;
                dl.Area.ForegroundColor = System.Drawing.Color.LightYellow;
                dl.Border.IsVisible = true;
                dl.Border.Color = System.Drawing.Color.DarkGray;
            }

            // ------------------------------------------------------------
            // 4. Save the workbook
            // ------------------------------------------------------------
            string outputPath = "BubbleChart_DataLabelsResized.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
