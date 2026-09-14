// Title: Validate custom slice colors in an Aspose.Cells pie chart by comparing ChartPoint.Area.ForegroundColor to expected RGB values in C#
// AI Prompts: Generate C# code that creates a pie chart with Aspose.Cells, assigns specific RGB colors to each slice via ChartPoint.Area.ForegroundColor, saves the workbook, reloads it, and verifies that each slice’s foreground color matches the predefined RGB array. | Write a C# routine that loads an existing Excel file containing a pie chart, iterates through its ChartPoint objects, compares their Area.ForegroundColor to an expected Color list, and logs any mismatches.
// Common Searches: asp.net verify pie chart slice colors Aspose.Cells C# | compare ChartPoint foreground color with expected RGB Aspose.Cells | how to check custom slice colors after saving workbook using Aspose.Cells .NET | C# read ChartPoint.Area.ForegroundColor from Excel pie chart
// Tags: Aspose.Cells pie chart custom slice colors | ChartPoint Area ForegroundColor verification C# | validate RGB colors in Excel chart Aspose.Cells | load workbook inspect chart point colors Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace VerifyCustomSliceColors
{
    // The example creates a workbook, adds data, and inserts a pie chart. It assigns specific RGB colors to each slice using ChartPoint.Area.ForegroundColor, saves the file, then reloads it. After loading, it iterates through the chart points, compares each slice's actual foreground color with the expected RGB values, outputs match/mismatch results, and saves a verified copy of the workbook.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart (Category and Value)
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

            // Set data range for the series
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Define expected colors for each slice (RGB)
            Color[] expectedColors = new Color[]
            {
                Color.FromArgb(255, 0, 0),      // Red for Apple
                Color.FromArgb(0, 255, 0),      // Green for Banana
                Color.FromArgb(0, 0, 255),      // Blue for Cherry
                Color.FromArgb(255, 255, 0)     // Yellow for Date
            };

            // Apply custom colors to each chart point (slice)
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                point.Area.ForegroundColor = expectedColors[i];
            }

            // Save the workbook (creation phase)
            string filePath = "CustomSliceColors.xlsx";
            workbook.Save(filePath);

            // ---------- Load the workbook and verify colors ----------
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Chart loadedChart = loadedSheet.Charts[chartIndex]; // same index as before

            // Verify each slice's foreground color against the expected RGB values
            for (int i = 0; i < loadedChart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = loadedChart.NSeries[0].Points[i];
                Color actual = point.Area.ForegroundColor;
                Color expected = expectedColors[i];

                bool match = actual.R == expected.R && actual.G == expected.G && actual.B == expected.B;

                if (match)
                {
                    Console.WriteLine($"Slice {i + 1} color matches expected RGB({expected.R},{expected.G},{expected.B}).");
                }
                else
                {
                    Console.WriteLine($"Slice {i + 1} color mismatch. Expected RGB({expected.R},{expected.G},{expected.B}) but got RGB({actual.R},{actual.G},{actual.B}).");
                }
            }

            // Optionally, save again after verification (no changes made)
            loadedWorkbook.Save("CustomSliceColors_Verified.xlsx");
        }
    }
}
