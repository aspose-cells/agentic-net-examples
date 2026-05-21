using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsPieChartCustomColors
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: categories, values, and importance (higher number = more important)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Importance");

            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["A5"].PutValue("Delta");
            sheet.Cells["A6"].PutValue("Epsilon");

            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["B5"].PutValue(15);
            sheet.Cells["B6"].PutValue(10);

            sheet.Cells["C2"].PutValue(5);   // most important
            sheet.Cells["C3"].PutValue(3);
            sheet.Cells["C4"].PutValue(4);
            sheet.Cells["C5"].PutValue(2);
            sheet.Cells["C6"].PutValue(1);   // least important

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the series (values) and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Define a color palette ordered by importance (most important first)
            Color[] importanceColors = new Color[]
            {
                Color.FromArgb(255, 0, 0),      // Red
                Color.FromArgb(255, 165, 0),    // Orange
                Color.FromArgb(255, 255, 0),    // Yellow
                Color.FromArgb(0, 128, 0),      // Green
                Color.FromArgb(0, 0, 255)       // Blue
            };

            // Iterate through each point (slice) and assign a custom color based on its importance value
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];

                // Read the importance from column C (same row as the point)
                int rowIndex = i + 2; // data starts at row 2
                int importance = Convert.ToInt32(sheet.Cells[$"C{rowIndex}"].Value);

                // Map importance (1..5) to color array index (0..4)
                int colorIndex = Math.Max(0, Math.Min(importanceColors.Length - 1, importance - 1));

                // Apply the color to the slice
                point.Area.ForegroundColor = importanceColors[colorIndex];
                point.Area.Formatting = FormattingType.Custom; // ensure custom formatting is used
            }

            // Save the workbook
            workbook.Save("CustomPieChartColors.xlsx");
        }
    }
}