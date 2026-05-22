using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", false);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply a built‑in chart style (styles range from 1 to 48)
                chart.Style = 10; // Example style number

                // Optionally adjust workbook theme colors
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(91, 155, 213)); // Light blue
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(237, 125, 49)); // Orange
                workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(165, 165, 165)); // Gray

                // Save the workbook
                string outputPath = "ChartWithBuiltInTheme.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}