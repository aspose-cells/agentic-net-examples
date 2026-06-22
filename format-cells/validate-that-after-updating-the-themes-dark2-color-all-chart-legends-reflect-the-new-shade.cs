using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemeValidation
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set data range for the series
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend uses a theme color (Accent2) for its font
            chart.Legend.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);

            // Save the workbook before changing the theme (optional step)
            workbook.Save("ThemeBeforeChange.xlsx");

            // ---------- Update the theme's Dark2 (mapped to Accent2) color ----------
            // Change Accent2 to a distinct color (e.g., Purple)
            workbook.SetThemeColor(ThemeColorType.Accent2, Color.Purple);

            // Save the workbook after the theme change
            workbook.Save("ThemeAfterChange.xlsx");

            // ---------- Load the saved workbook to validate ----------
            Workbook loadedWb = new Workbook("ThemeAfterChange.xlsx");
            Worksheet loadedSheet = loadedWb.Worksheets[0];
            Chart loadedChart = loadedSheet.Charts[chartIdx];

            // Retrieve the legend font's ThemeColor
            ThemeColor legendThemeColor = loadedChart.Legend.Font.ThemeColor;

            // Validate that the legend still references Accent2 (the updated theme color)
            bool isAccent2 = legendThemeColor != null && legendThemeColor.ColorType == ThemeColorType.Accent2;

            Console.WriteLine(isAccent2
                ? "Validation passed: Legend font uses the updated Accent2 theme color."
                : "Validation failed: Legend font does not use the expected theme color.");

            // Optionally, save the final workbook
            loadedWb.Save("ThemeValidationResult.xlsx");
        }
    }
}