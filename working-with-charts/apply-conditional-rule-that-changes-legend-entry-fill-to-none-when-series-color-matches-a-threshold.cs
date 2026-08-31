// Title: How to conditionally remove chart legend entry fill in Aspose.Cells for .NET when the series color exceeds a red‑component threshold
// AI Prompts: Write C# code using Aspose.Cells that checks the foreground color of a chart series and sets the series' LegendEntry.IsTextNoFill property to true when the red component is greater than a specified value. | Outline the code required to set a legend entry's IsTextNoFill flag when a series color surpasses a given RGB threshold in Aspose.Cells.
// Common Searches: Aspose.Cells C# hide legend entry fill based on series color | column chart legend fill removal when series color red component exceeds threshold using Aspose.Cells | programmatically control IsTextNoFill for chart legend entry in .NET
// Tags: conditional legend formatting Aspose.Cells | chart series color threshold .NET | set legend entry no fill C# | column chart legend IsTextNoFill Aspose.Cells | chart legend conditional rule Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendConditionalDemo
{
    // The example creates a workbook, adds a column chart with sample data, assigns a custom foreground color to the series, defines a red‑component threshold, and then checks the series color. If the red value exceeds the threshold, it sets the legend entry's IsTextNoFill property to true, otherwise it keeps the default fill, and finally saves the workbook as LegendConditionalDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            // Categories
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");

            // Values for the series
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(80);
            worksheet.Cells["B4"].PutValue(150);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Bind data to the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Set a specific color for the first (and only) series
            // This color will be used to decide whether to remove the legend fill
            chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(200, 100, 50); // Example color

            // Define a threshold color (for demonstration we use a simple RGB comparison)
            // Here we consider the threshold as any color with Red component > 180
            Color thresholdColor = Color.FromArgb(180, 0, 0);

            // Access the legend entry associated with the series
            LegendEntry legendEntry = chart.NSeries[0].LegendEntry;

            // Check if the series color meets the threshold condition
            Color seriesColor = chart.NSeries[0].Area.ForegroundColor;
            if (seriesColor.R > thresholdColor.R) // Simple threshold logic
            {
                // When condition is met, set the legend entry text to have no fill
                legendEntry.IsTextNoFill = true;
            }
            else
            {
                // Otherwise ensure normal fill
                legendEntry.IsTextNoFill = false;
            }

            // Save the workbook
            workbook.Save("LegendConditionalDemo.xlsx");
        }
    }
}
