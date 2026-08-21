// Title: Auto‑Fit Chart Size in Aspose.Cells (C#) Using GetActualSize
// Description: Shows how to create a workbook, add a column chart, recalculate its layout, retrieve the exact pixel dimensions with Chart.GetActualSize(), and apply those values to ChartObject.Width and Height to avoid clipping. Includes optional AutoScaleFont for the title.
// Keywords: Aspose.Cells | C# | auto fit chart | GetActualSize | ChartObject dimensions | prevent chart clipping | chart resizing | Excel chart size
// Common Searches: Aspose.Cells resize chart to fit data | GetActualSize chart dimensions C# | auto fit chart size Aspose | prevent chart clipping Aspose.Cells | set ChartObject width height programmatically
// Developer Intent: Programmatically set a chart's width and height to the exact size Excel would allocate, ensuring the visual content is fully displayed.
// Use Cases: Create a column chart from a data range and automatically adjust its dimensions before saving the workbook. | Loop through multiple charts in a report, applying the same auto‑fit logic to maintain consistent appearance. | Refresh an existing chart after changing its data source by recalculating layout and applying the new actual size.
// AI Prompts: Generate C# code that adds a line chart with Aspose.Cells, calls Calculate(), retrieves the actual size, and sets ChartObject.Width and Height. | Write a reusable method that accepts a Chart object and resizes it to its optimal dimensions using GetActualSize, with error handling. | Explain how to enable AutoScaleFont for a chart title after resizing the chart in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoFitDemo
{
    // Shows how to create a workbook, add a column chart, recalculate its layout, retrieve the exact pixel dimensions with Chart.GetActualSize(), and apply those values to ChartObject.Width and Height to avoid clipping. Includes optional AutoScaleFont for the title.
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

            // Add a column chart (initial size is arbitrary)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Force the chart to recalculate its layout based on the data
            chart.Calculate();

            // Retrieve the actual size (in pixels) that Excel would use for this chart
            int[] actualSize = chart.GetActualSize(); // [0] = width, [1] = height

            // Apply the calculated size to the chart object to avoid clipping
            chart.ChartObject.Width = actualSize[0];
            chart.ChartObject.Height = actualSize[1];

            // Optional: ensure the title scales with the new size
            chart.Title.AutoScaleFont = true;

            // Save the workbook
            workbook.Save("ChartAutoFitDemo.xlsx");
        }
    }
}
