using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class VerifyChartSliceColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data for a pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["A4"].PutValue("Cherry");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(45);
        sheet.Cells["B4"].PutValue(25);

        // Add a pie chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Set the series data and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Define custom colors for each slice (Red, Green, Blue)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 0, 0),   // Red
            Color.FromArgb(255, 0, 255, 0),   // Green
            Color.FromArgb(255, 0, 0, 255)    // Blue
        };

        // Apply the custom colors to each point in the series
        for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
        {
            chart.NSeries[0].Points[i].Area.ForegroundColor = customColors[i];
        }

        // Verify that the colors were set correctly
        for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
        {
            Color actual = chart.NSeries[0].Points[i].Area.ForegroundColor;
            Color expected = customColors[i];
            bool match = actual.ToArgb() == expected.ToArgb();
            Console.WriteLine($"Slice {i + 1}: Expected RGB({expected.R},{expected.G},{expected.B}) " +
                              $"Actual RGB({actual.R},{actual.G},{actual.B}) - Match: {match}");
        }

        // Save the workbook
        workbook.Save("PieChartCustomSliceColors.xlsx");
    }
}