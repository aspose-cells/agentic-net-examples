// Title: Resize data label font size for a bubble chart after scaling bubble sizes with Aspose.Cells for .NET
// AI Prompts: Create a new Excel workbook, add a bubble chart, multiply the size values by a scaling factor, recalculate the chart, and set the data label font size proportionally using Aspose.Cells in C#. | Update an existing bubble chart's data label font size to reflect a custom scaling factor applied to the bubble size series in a .NET application. | Write C# code that reads bubble size data, applies a multiplier, refreshes the chart, and adjusts the data label font size before saving the workbook with Aspose.Cells.
// Common Searches: how to change bubble chart data label font size with Aspose.Cells C# | scale bubble sizes and automatically resize labels in Excel using .NET | Aspose.Cells example for adjusting data label size after increasing bubble size | C# code to recalculate bubble chart after modifying size values | programmatically increase bubble chart markers and update label font in Aspose.Cells
// Tags: bubble chart data label font resizing Aspose.Cells | scale bubble sizes chart .NET | adjust chart series data labels C# | Aspose.Cells bubble chart scaling example | programmatic Excel chart label customization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a bubble chart, scaling the bubble size values by a factor, recalculating the chart, and proportionally resizing the data label font before saving the workbook using Aspose.Cells for .NET.
class ResizeBubbleChartDataLabels
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a bubble chart
            // Column A: X values, Column B: Y values, Column C: Bubble sizes
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Size");

            double[] xValues = { 1, 2, 3, 4, 5 };
            double[] yValues = { 2, 4, 1, 3, 5 };
            double[] sizes   = { 10, 20, 30, 40, 50 }; // initial bubble sizes

            for (int i = 0; i < xValues.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(xValues[i]); // A column
                sheet.Cells[i + 1, 1].PutValue(yValues[i]); // B column
                sheet.Cells[i + 1, 2].PutValue(sizes[i]);   // C column
            }

            // Add a bubble chart
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 7, 0, 27, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series and set its data range (X, Y, Size)
            int seriesIndex = chart.NSeries.Add("A2:C6", true);
            Series series = chart.NSeries[seriesIndex];

            // Show data labels for the series
            series.DataLabels.ShowValue = true;

            // Increase bubble sizes to test scaling (multiply by a factor)
            double scaleFactor = 3.0;
            for (int i = 0; i < sizes.Length; i++)
            {
                double newSize = sizes[i] * scaleFactor;
                sheet.Cells[i + 1, 2].PutValue(newSize);
            }

            // Refresh the chart to reflect new bubble sizes
            chart.Calculate();

            // Resize data label font size proportionally to bubble scaling
            foreach (Series ser in chart.NSeries)
            {
                if (ser.DataLabels != null)
                {
                    ser.DataLabels.Font.Size = (int)(8 * scaleFactor); // base size 8pt
                }
            }

            // Save the workbook
            string outputPath = "BubbleChart_WithResizedDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
