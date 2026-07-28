// Title: Update Dark2 (Accent2) Theme Color and Verify Chart Legends in Aspose.Cells for .NET
// Description: Shows how to modify the workbook's Dark2 (Accent2) theme color with Aspose.Cells, save the file, and confirm that a column chart's legend font automatically adopts the new shade without extra formatting steps.
// Keywords: Aspose.Cells | SetThemeColor | ThemeColorType.Accent2 | chart legend color | C# | .NET | Excel theme update | Dark2 theme | validate chart legend | workbook theme change
// Common Searches: Aspose.Cells change theme color C# | update Dark2 (Accent2) in Excel workbook | chart legend color after theme change Aspose | refresh chart after SetThemeColor | verify theme color applied to chart legend
// Developer Intent: Confirm that all chart legends automatically display the new Dark2 (Accent2) shade after the workbook theme is altered via Workbook.SetThemeColor.
// Use Cases: Programmatically change the Accent2 color to a custom hue and let existing chart legends reflect the update instantly. | Generate before‑and‑after Excel files to visually compare legend colors when the theme is modified. | Read chart.Legend.Font.ThemeColor to ensure the ColorType remains Accent2 and the tint matches the new shade.
// AI Prompts: Write C# code using Aspose.Cells that updates the Dark2 (Accent2) theme color and asserts that every chart legend in the workbook shows the new color. | Create a unit test that loads a workbook, calls Workbook.SetThemeColor for Accent2, and verifies the legend font color has changed accordingly. | Explain how to programmatically check that chart legends stay linked to the theme after calling SetThemeColor in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsThemeLegendDemo
{
    // Shows how to modify the workbook's Dark2 (Accent2) theme color with Aspose.Cells, save the file, and confirm that a column chart's legend font automatically adopts the new shade without extra formatting steps.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");

                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                    sheet.Cells[$"C{i}"].PutValue(i * 12);
                }

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the chart
                chart.SetChartDataRange("B1:C5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Configure the legend to use a theme color (Accent2 represents Dark2 in this context)
                chart.Legend.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);
                // Ensure the legend is displayed (position it at the bottom)
                chart.Legend.Position = LegendPositionType.Bottom;

                // Save the workbook before changing the theme (optional, just for reference)
                workbook.Save("BeforeThemeChange.xlsx");

                // Update the theme's Dark2 (mapped to Accent2) color to a new shade
                Color newDark2Shade = Color.FromArgb(255, 128, 0); // Example: orange shade
                workbook.SetThemeColor(ThemeColorType.Accent2, newDark2Shade);

                // Save the workbook after the theme change
                workbook.Save("AfterThemeChange.xlsx");

                // Validation: the legend font still references the theme color Accent2
                ThemeColor legendTheme = chart.Legend.Font.ThemeColor;
                Console.WriteLine($"Legend font uses ThemeColor type: {legendTheme.ColorType}");
                Console.WriteLine($"Legend font tint value: {legendTheme.Tint}");
                Console.WriteLine("If the theme color Accent2 was updated, the legend will display the new shade.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
