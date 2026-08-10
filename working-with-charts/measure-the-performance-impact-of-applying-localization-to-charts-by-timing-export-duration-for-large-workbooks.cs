// Title: Chart Localization Performance Test – Export Large Workbook to PDF with Aspose.Cells for .NET
// Description: C# sample that builds a 10,000‑row workbook with a column chart, applies a custom ChartGlobalizationSettings to localize axis unit names, and measures PDF export time with and without the localization using Stopwatch, reporting the elapsed milliseconds for each run.
// Keywords: Aspose.Cells | C# chart localization | performance benchmarking | PDF export timing | large workbook | ChartGlobalizationSettings | globalization impact | export speed test
// Common Searches: Aspose.Cells chart localization performance | measure PDF export time for large workbook | benchmark Aspose.Cells chart globalization | how long does chart localization add to export | timing Aspose.Cells PDF save with custom chart settings
// Developer Intent: Find out how custom chart globalization settings affect the time required to export a large workbook to PDF using Aspose.Cells.
// Use Cases: Compare export duration of a 10k‑row workbook with and without custom ChartGlobalizationSettings. | Validate that applying localized axis labels does not cause unacceptable slowdown in PDF generation. | Profile Aspose.Cells export performance for datasets that include charts and globalization overrides.
// AI Prompts: Create C# code that logs PDF and XPS export times for a 20,000‑row workbook with and without ChartGlobalizationSettings. | Suggest optimization techniques to minimize the overhead introduced by custom chart localization in Aspose.Cells. | Write a unit test in C# that asserts the export time difference between localized and non‑localized charts stays under a defined threshold.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPerformanceDemo
{
    // Custom globalization settings for charts (e.g., localized axis unit names)
    // C# sample that builds a 10,000‑row workbook with a column chart, applies a custom ChartGlobalizationSettings to localize axis unit names, and measures PDF export time with and without the localization using Stopwatch, reporting the elapsed milliseconds for each run.
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Example: localize display unit names to Chinese
            return type switch
            {
                DisplayUnitType.Hundreds => "百",
                DisplayUnitType.Thousands => "千",
                DisplayUnitType.TenThousands => "万",
                _ => base.GetAxisUnitName(type),
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a large workbook with sample data and a chart
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate a large dataset (e.g., 10000 rows)
            int rows = 10000;
            cells[0, 0].PutValue("Category");
            cells[0, 1].PutValue("Value");
            for (int i = 1; i <= rows; i++)
            {
                cells[i, 0].PutValue($"Item {i}");
                cells[i, 1].PutValue(i % 1000 + 1); // some varying values
            }

            // Add a column chart covering the whole data range
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 30, 10);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add($"B2:B{rows + 1}", true);          // values
            chart.NSeries.CategoryData = $"A2:A{rows + 1}";     // categories
            chart.Title.Text = "Large Data Chart";

            // -----------------------------------------------------------------
            // 2. Apply localization (globalization) to the chart
            // -----------------------------------------------------------------
            wb.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new CustomChartGlobalizationSettings()
            };

            // Force chart to recalculate layout after changing globalization
            chart.Calculate();

            // -----------------------------------------------------------------
            // 3. Measure export duration (e.g., to PDF) without localization
            // -----------------------------------------------------------------
            // Clone workbook to have a version without localization for comparison
            Workbook wbNoLocalization = new Workbook();
            wbNoLocalization.Copy(wb);
            // Remove localization from the clone
            wbNoLocalization.Settings.GlobalizationSettings = new GlobalizationSettings();

            // Export without localization
            Stopwatch sw = new Stopwatch();
            sw.Start();
            wbNoLocalization.Save("LargeWorkbook_NoLocalization.pdf", SaveFormat.Pdf);
            sw.Stop();
            Console.WriteLine($"Export without localization: {sw.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // 4. Measure export duration with localization applied
            // -----------------------------------------------------------------
            sw.Restart();
            wb.Save("LargeWorkbook_WithLocalization.pdf", SaveFormat.Pdf);
            sw.Stop();
            Console.WriteLine($"Export with localization: {sw.ElapsedMilliseconds} ms");

            // Cleanup
            wb.Dispose();
            wbNoLocalization.Dispose();
        }
    }
}
