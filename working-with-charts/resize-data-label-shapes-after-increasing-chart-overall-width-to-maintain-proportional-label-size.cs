// Title: Resize Chart Data Label Shapes Proportionally After Expanding Width – Aspose.Cells for .NET
// Description: Demonstrates how to double a column chart's width, compute a scaling factor, disable auto‑fit, and set WidthPixel/HeightPixel on data labels so they stay proportional before saving the workbook.
// Keywords: Aspose.Cells chart resize | data label shape scaling | WidthPixel HeightPixel Aspose.Cells | disable auto‑fit chart labels | C# chart label proportional size
// Common Searches: Aspose.Cells resize data label after chart width change | set data label WidthPixel HeightPixel in .NET | disable auto fit for chart labels Aspose.Cells | calculate scaling factor for chart labels C# | proportional label size when expanding chart Aspose
// Developer Intent: Adjust the dimensions of chart data label shapes so they remain proportional after the chart’s width is increased.
// Use Cases: Enlarge a dashboard chart while keeping label readability consistent. | Prepare a wide printable chart without distorting label appearance. | Batch‑process multiple charts to maintain uniform label sizing after scaling.
// AI Prompts: Generate C# code using Aspose.Cells that resizes data label shapes proportionally after changing a chart's width. | Explain how to turn off auto‑fit and assign explicit WidthPixel and HeightPixel values to chart data labels in Aspose.Cells. | Show the formula for computing a width‑based scaling factor and applying it to data label dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLabelResizeDemo
{
    // Demonstrates how to double a column chart's width, compute a scaling factor, disable auto‑fit, and set WidthPixel/HeightPixel on data labels so they stay proportional before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Calculate the chart so that shape dimensions become available
            chart.Calculate();

            // Store original chart width and original data label size (in pixels)
            int originalChartWidth = chart.ChartObject.Width;
            int originalLabelWidth = series.DataLabels.WidthPixel;
            int originalLabelHeight = series.DataLabels.HeightPixel;

            // Increase the chart width (e.g., double it)
            chart.ChartObject.Width = originalChartWidth * 2;

            // Recalculate after resizing the chart
            chart.Calculate();

            // Compute scaling factor based on chart width change
            double widthScale = (double)chart.ChartObject.Width / originalChartWidth;

            // Disable auto‑fit so we can set explicit dimensions
            series.DataLabels.IsResizeShapeToFitText = false;

            // Apply proportional size to the data label shape
            series.DataLabels.WidthPixel = (int)(originalLabelWidth * widthScale);
            series.DataLabels.HeightPixel = (int)(originalLabelHeight * widthScale);

            // Save the workbook
            workbook.Save("ChartLabelResizeDemo.xlsx");
        }
    }
}
