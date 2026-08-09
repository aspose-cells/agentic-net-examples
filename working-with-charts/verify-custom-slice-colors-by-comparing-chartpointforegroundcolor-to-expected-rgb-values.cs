// Title: Verify Pie Chart Slice Colors Using ChartPoint.ForegroundColor in Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a pie chart with four categories, assigns specific RGB colors to each slice via ChartPoint.Area.ForegroundColor, compares the actual colors with the expected values, logs the results, and saves the file. Ideal for developers who need to programmatically confirm chart color schemes in Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | pie chart | slice color | ChartPoint | ForegroundColor | RGB validation | chart color verification | automated testing | chart customization
// Common Searches: how to set pie chart slice colors Aspose.Cells C# | verify chart point color Aspose.Cells .NET | compare ChartPoint.ForegroundColor with expected RGB | Aspose.Cells color validation for charts | programmatically check pie chart colors in C#
// Developer Intent: Confirm that each slice of a generated pie chart matches the predefined RGB color values.
// Use Cases: Enforce brand‑compliant color palettes in automatically generated reports. | Run regression tests to ensure chart colors remain consistent after code changes. | Dynamically assign and validate slice colors based on data‑driven rules.
// AI Prompts: Write C# code with Aspose.Cells that sets custom RGB colors for each slice of a pie chart and then verifies the colors by comparing ChartPoint.Area.ForegroundColor to a predefined Color array. | Create a method that iterates through ChartPoint objects, extracts their ForegroundColor, and returns the indices of any mismatched RGB values. | Explain how to log detailed color comparison results for chart points, handling cases where the number of points differs from the expected colors array.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds a pie chart with four categories, assigns specific RGB colors to each slice via ChartPoint.Area.ForegroundColor, compares the actual colors with the expected values, logs the results, and saves the file. Ideal for developers who need to programmatically confirm chart color schemes in Aspose.Cells.
    public class VerifyChartPointSliceColors
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pie chart (categories and values)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["A5"].PutValue("Date");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(25);
                sheet.Cells["B5"].PutValue(25);

                // Add a pie chart
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Define custom slice colors (RGB)
                Color[] expectedColors = new Color[]
                {
                    Color.FromArgb(255, 255, 0, 0),   // Red
                    Color.FromArgb(255, 0, 255, 0),   // Green
                    Color.FromArgb(255, 0, 0, 255),   // Blue
                    Color.FromArgb(255, 255, 165, 0)  // Orange
                };

                // Apply custom colors to each point (slice) in the first series
                Series series = chart.NSeries[0];
                for (int i = 0; i < series.Points.Count && i < expectedColors.Length; i++)
                {
                    ChartPoint point = series.Points[i];
                    point.Area.ForegroundColor = expectedColors[i];
                }

                // Verify that each point's foreground color matches the expected RGB value
                for (int i = 0; i < series.Points.Count && i < expectedColors.Length; i++)
                {
                    ChartPoint point = series.Points[i];
                    Color actualColor = point.Area.ForegroundColor;
                    Color expectedColor = expectedColors[i];

                    bool match = actualColor.ToArgb() == expectedColor.ToArgb();
                    Console.WriteLine($"Slice {i + 1}: Expected RGB({expectedColor.R},{expectedColor.G},{expectedColor.B}) " +
                                      $"- Actual RGB({actualColor.R},{actualColor.G},{actualColor.B}) " +
                                      $"=> Match: {match}");
                }

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerifyChartPointSliceColors.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyChartPointSliceColors.Run();
        }
    }
}
