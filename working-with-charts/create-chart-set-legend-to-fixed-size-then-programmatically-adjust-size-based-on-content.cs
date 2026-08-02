// Title: Aspose.Cells C# – Create a Column Chart and Dynamically Resize Its Legend
// Description: This example shows how to build a workbook, add a column chart, turn off automatic legend sizing, calculate the chart to obtain legend labels, compute the required pixel width and height based on label length and count, and then apply those dimensions before saving the file.
// Keywords: Aspose.Cells chart legend resize | C# dynamic legend size | GetLegendLabels Aspose.Cells | chart legend WidthPixel HeightPixel | disable automatic legend sizing | column chart Aspose.Cells .NET
// Common Searches: Aspose.Cells resize chart legend programmatically | C# calculate legend width from label length | How to get legend labels in Aspose.Cells | Set fixed legend size then adjust in .NET | Dynamic legend dimensions for Excel chart
// Developer Intent: Adjust a chart legend so its width and height automatically fit the longest label and total entries.
// Use Cases: Generate Excel reports where the bottom legend expands to accommodate long series names without truncation. | Create dashboards that automatically increase legend height to prevent overlap when many series are displayed. | Build reusable chart components that start with a placeholder legend size and resize after data is known.
// AI Prompts: Write C# code using Aspose.Cells to create a column chart and set the legend size based on the maximum label length returned by GetLegendLabels. | Provide a helper method that accepts a Chart object, calculates required WidthPixel and HeightPixel for the legend, and applies the values. | Explain the steps to disable automatic legend sizing, retrieve legend labels, and compute pixel dimensions for dynamic resizing in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDynamicSize
{
    // This example shows how to build a workbook, add a column chart, turn off automatic legend sizing, calculate the chart to obtain legend labels, compute the required pixel width and height based on label length and count, and then apply those dimensions before saving the file.
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);

            // Set legend to a fixed size initially
            chart.Legend.IsAutomaticSize = false;               // disable automatic sizing
            chart.Legend.Position = LegendPositionType.Bottom;  // place legend at bottom
            chart.Legend.WidthPixel = 200;                      // provisional width
            chart.Legend.HeightPixel = 50;                      // provisional height

            // Calculate the chart so that legend labels are generated
            chart.Calculate();

            // Retrieve legend labels after calculation
            string[] legendLabels = chart.Legend.GetLegendLabels();

            // Determine the maximum label length
            int maxLabelLength = 0;
            foreach (string label in legendLabels)
            {
                if (label != null && label.Length > maxLabelLength)
                    maxLabelLength = label.Length;
            }

            // Approximate required width (7 pixels per character + padding)
            int requiredWidth = maxLabelLength * 7 + 20;

            // Approximate required height (15 pixels per entry + padding)
            int requiredHeight = legendLabels.Length * 15 + 10;

            // Adjust legend size based on content
            chart.Legend.WidthPixel = requiredWidth;
            chart.Legend.HeightPixel = requiredHeight;

            // Save the workbook
            workbook.Save("ChartLegendDynamicSize.xlsx");
        }
    }
}
