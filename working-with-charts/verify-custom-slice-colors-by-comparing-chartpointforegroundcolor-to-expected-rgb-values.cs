using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace VerifyChartSliceColors
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
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

            // Define custom colors for each slice (RGB)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 0, 0),      // Red for Apple
                Color.FromArgb(255, 255, 0),    // Yellow for Banana
                Color.FromArgb(255, 0, 255),    // Magenta for Cherry
                Color.FromArgb(0, 128, 0)       // Green for Date
            };

            // Apply custom colors to each chart point
            for (int i = 0; i < customColors.Length; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                point.Area.ForegroundColor = customColors[i];
            }

            // Verify that the applied colors match the expected RGB values
            bool allMatch = true;
            for (int i = 0; i < customColors.Length; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                Color actual = point.Area.ForegroundColor;
                Color expected = customColors[i];

                bool match = actual.ToArgb() == expected.ToArgb();
                Console.WriteLine($"Slice {i + 1}: Expected RGB({expected.R},{expected.G},{expected.B}) " +
                                  $"- Actual RGB({actual.R},{actual.G},{actual.B}) => " +
                                  (match ? "Match" : "Mismatch"));

                if (!match) allMatch = false;
            }

            Console.WriteLine(allMatch
                ? "All slice colors match the expected values."
                : "One or more slice colors do not match the expected values.");

            // Save the workbook
            workbook.Save("PieChartWithCustomSliceColors.xlsx");
        }
    }
}