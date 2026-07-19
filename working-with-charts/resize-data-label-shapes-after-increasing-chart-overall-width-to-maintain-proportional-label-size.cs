// Title: C# – Resize Aspose.Cells Chart Data Labels Proportionally After Expanding Chart Width
// Description: Demonstrates how to double an Aspose.Cells column chart's width in .NET while preserving the original size of data‑label shapes. The example records the label's WidthRatioToChart, enlarges the chart, recalculates, disables automatic shape resizing, and reapplies the stored ratio so labels stay proportional before saving the workbook.
// Keywords: Aspose.Cells chart resize C# | data label WidthRatioToChart | IsResizeShapeToFitText property | maintain label size after chart width change | .NET chart data labels scaling | Aspose.Cells chart.Calculate usage | proportional label resizing
// Common Searches: Aspose.Cells keep data label size when chart width changes | C# WidthRatioToChart example | how to prevent label shape from stretching in Aspose.Cells | resize chart and preserve label dimensions .NET | chart.Calculate after resizing Aspose.Cells
// Developer Intent: Enlarge an Aspose.Cells chart while keeping its data‑label shapes proportionally sized.
// Use Cases: Doubling a chart's width for presentation layouts without distorting label fonts. | Applying the same proportional label size to multiple series after a chart resize. | Creating responsive Excel reports where chart dimensions vary but label readability must stay consistent.
// AI Prompts: Write C# code that expands an Aspose.Cells chart and retains data‑label proportions using WidthRatioToChart and IsResizeShapeToFitText. | Explain the relationship between WidthRatioToChart, HeightRatioToChart, and IsResizeShapeToFitText when adjusting chart size in Aspose.Cells. | Provide a step‑by‑step tutorial for resizing a chart and updating data‑label shape ratios for all series in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLabelResizeDemo
{
    // Demonstrates how to double an Aspose.Cells column chart's width in .NET while preserving the original size of data‑label shapes. The example records the label's WidthRatioToChart, enlarges the chart, recalculates, disables automatic shape resizing, and reapplies the stored ratio so labels stay proportional before saving the workbook.
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

            // Calculate the chart so that shape dimensions become available
            chart.Calculate();

            // Store the original width ratio of the data label shape (relative to chart width)
            double originalLabelWidthRatio = series.DataLabels.WidthRatioToChart;

            // Increase the overall chart width (e.g., double it)
            chart.ChartObject.Width = chart.ChartObject.Width * 2;

            // Re‑calculate after resizing the chart
            chart.Calculate();

            // Adjust each data label to keep the same proportional size
            foreach (Series s in chart.NSeries)
            {
                // Disable automatic resizing of the shape to fit the text
                s.DataLabels.IsResizeShapeToFitText = false;

                // Apply the original width ratio so the label scales with the chart size
                s.DataLabels.WidthRatioToChart = originalLabelWidthRatio;

                // Optionally, you can also keep the height proportional
                // s.DataLabels.HeightRatioToChart = originalLabelHeightRatio;
            }

            // Save the workbook
            workbook.Save("ChartLabelResizeDemo.xlsx");
        }
    }
}
