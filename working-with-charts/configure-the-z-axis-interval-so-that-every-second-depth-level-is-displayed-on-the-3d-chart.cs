// Title: Configure Z‑Axis Interval to Show Every Other Depth Level in a 3‑D Column Chart – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, insert a 3‑D column chart, and set the SeriesAxis.TickMarkSpacing property to 2 so the Z (depth) axis displays a tick for every second level. Includes optional rotation, elevation, and perspective adjustments before saving the file.
// Keywords: Aspose.Cells Z axis interval | SeriesAxis.TickMarkSpacing C# | 3D column chart depth axis | display alternate depth levels Aspose | configure Z axis Aspose.Cells .NET | 3D chart spacing series axis
// Common Searches: Aspose.Cells set Z axis tick spacing | how to show every other depth level in 3D chart | SeriesAxis.TickMarkSpacing example .NET | configure depth axis interval Aspose.Cells | 3D column chart spacing series axis C#
// Developer Intent: Set the Z (depth) axis of a 3‑D chart to display ticks for every second series/category.
// Use Cases: Reduce clutter in dense 3‑D column charts by skipping alternate depth levels. | Highlight alternating series for clearer visual comparison. | Improve readability when a chart contains many categories or series.
// AI Prompts: Write C# code with Aspose.Cells that creates a 3‑D bar chart and sets SeriesAxis.TickMarkSpacing to 3. | Explain the effect of SeriesAxis.TickMarkSpacing on the depth axis of a 3‑D chart and how to revert to the default value. | Provide step‑by‑step instructions to adjust rotation, elevation, and perspective after configuring Z‑axis tick spacing in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, insert a 3‑D column chart, and set the SeriesAxis.TickMarkSpacing property to 2 so the Z (depth) axis displays a tick for every second level. Includes optional rotation, elevation, and perspective adjustments before saving the file.
    public class ConfigureZAxisIntervalDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (categories and two series)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Level 1");
                worksheet.Cells["A3"].PutValue("Level 2");
                worksheet.Cells["A4"].PutValue("Level 3");
                worksheet.Cells["A5"].PutValue("Level 4");
                worksheet.Cells["A6"].PutValue("Level 5");

                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);
                worksheet.Cells["B5"].PutValue(40);
                worksheet.Cells["B6"].PutValue(50);

                worksheet.Cells["C1"].PutValue("Series 2");
                worksheet.Cells["C2"].PutValue(15);
                worksheet.Cells["C3"].PutValue(25);
                worksheet.Cells["C4"].PutValue(35);
                worksheet.Cells["C5"].PutValue(45);
                worksheet.Cells["C6"].PutValue(55);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 7, 0, 25, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:C6", true);          // Values
                chart.NSeries.CategoryData = "A2:A6";     // Categories

                // Configure the Z (depth/series) axis to show every second depth level
                // In a 3‑D chart the depth axis is the SeriesAxis; setting TickMarkSpacing to 2
                // displays a tick (and thus a depth level) for every second series/category.
                chart.SeriesAxis.TickMarkSpacing = 2;

                // Optional: adjust other 3‑D view properties for better visibility
                chart.RotationAngle = 30;
                chart.Elevation = 20;
                chart.Perspective = 40;

                // Save the workbook
                workbook.Save("ConfigureZAxisIntervalDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConfigureZAxisIntervalDemo.Run();
        }
    }
}
