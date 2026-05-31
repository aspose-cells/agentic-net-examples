using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPerformanceDemo
{
    // Custom globalization settings for charts (e.g., change axis unit names)
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Example: prepend a custom prefix to the default unit name
            string defaultName = base.GetAxisUnitName(type);
            return "Loc-" + defaultName;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Generate a large dataset (e.g., 10,000 rows, 5 columns)
            int rows = 10000;
            int cols = 5;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(r * cols + c + 1);
                }
            }

            // Add several charts that reference the data
            int chartCount = 10;
            for (int i = 0; i < chartCount; i++)
            {
                // Position charts vertically stacked
                int topRow = i * 15;
                int chartIndex = sheet.Charts.Add(ChartType.Column, topRow, 0, topRow + 14, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                string dataRange = $"B1:B{rows}";
                string categoryRange = $"A1:A{rows}";
                chart.NSeries.Add(dataRange, true);
                chart.NSeries.CategoryData = categoryRange;

                // Enable a display unit to see globalization effect
                chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
                chart.ValueAxis.IsDisplayUnitLabelShown = true;
            }

            // -----------------------------------------------------------------
            // Measure export time without localization (default settings)
            // -----------------------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Export to PDF (lifecycle rule: save)
            workbook.Save("Export_Default.pdf", SaveFormat.Pdf);

            sw.Stop();
            Console.WriteLine($"Export without localization: {sw.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Apply localization to charts via globalization settings
            // -----------------------------------------------------------------
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new CustomChartGlobalizationSettings()
            };

            // Recalculate charts to apply new globalization (optional but ensures changes)
            foreach (Chart chart in sheet.Charts)
            {
                chart.Calculate();
            }

            // Measure export time with localization
            sw.Restart();

            workbook.Save("Export_Localized.pdf", SaveFormat.Pdf);

            sw.Stop();
            Console.WriteLine($"Export with localization: {sw.ElapsedMilliseconds} ms");

            // Clean up
            workbook.Dispose();
        }
    }
}