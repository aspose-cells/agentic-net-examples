// Title: Manually resize bubble chart data label shapes after increasing bubble scale with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that sets a bubble chart's BubbleScale to 200% and assigns a fixed width and height to each data label shape. | Show how to turn off automatic resizing of data label shapes and specify custom dimensions for bubble chart points in a .NET workbook.
// Common Searches: Aspose.Cells C# set bubble chart label width and height | increase bubble scale while keeping data label size constant Aspose.Cells | prevent data label shape auto‑sizing in bubble chart using Aspose.Cells .NET | programmatically adjust bubble chart data label dimensions Aspose.Cells | how to change bubble chart data label size after scaling in C#
// Tags: bubble chart data label custom size Aspose.Cells | set bubble scale Aspose.Cells C# | data label auto‑resize off Aspose.Cells | manual data label dimensions .NET chart | Aspose.Cells bubble chart label resizing

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a bubble chart with sample X, Y, and size data, sets the BubbleScale to 200%, enables data labels to show values and bubble sizes, disables automatic shape resizing for each point, assigns a fixed width of 80 and height of 30 points to the data label shapes, recalculates the chart, and saves the file as BubbleChartDataLabelResize.xlsx.
class ResizeDataLabelShapesBubbleChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the bubble chart
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["C1"].PutValue("Size");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(5);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["C3"].PutValue(10);
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["C4"].PutValue(15);

        // Add a bubble chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series and assign X, Y and bubble size ranges
        int seriesIndex = chart.NSeries.Add("B2:B4", true);
        chart.NSeries[seriesIndex].XValues = "A2:A4";
        chart.NSeries[seriesIndex].BubbleSizes = "C2:C4";

        // Increase bubble size scaling to 200% to test label scaling
        chart.NSeries[seriesIndex].BubbleScale = 200;

        // Enable data labels and display bubble size values
        Series series = chart.NSeries[seriesIndex];
        series.DataLabels.ShowValue = true;
        series.DataLabels.ShowBubbleSize = true;

        // Resize each data label shape manually
        foreach (ChartPoint point in series.Points)
        {
            // Disable automatic shape resizing to fit text
            point.DataLabels.IsResizeShapeToFitText = false;

            // Set custom dimensions for the data label shape (units are points)
            point.DataLabels.Width = 80;
            point.DataLabels.Height = 30;
        }

        // Recalculate the chart to apply the changes
        chart.Calculate();

        // Save the workbook with the modified chart
        workbook.Save("BubbleChartDataLabelResize.xlsx");
    }
}
