// Title: Benchmark Chart Localization Performance on Large Workbooks with Aspose.Cells for .NET
// Description: Creates a workbook with 10,000 rows, adds a column chart, measures PDF export time, then applies a custom ChartGlobalizationSettings that localizes axis unit names, recalculates the chart, and measures the export time again. The two durations are printed for direct comparison.
// Keywords: Aspose.Cells | C# chart localization | ChartGlobalizationSettings | PDF export performance | large workbook benchmark | measure export time | custom axis unit names | Aspose.Cells performance testing
// Common Searches: Aspose.Cells chart localization performance test | How to measure PDF export time with localized charts in .NET | Benchmark Aspose.Cells export with custom ChartGlobalizationSettings | Performance impact of chart globalization in large Excel files | Timing Aspose.Cells PDF generation for 10,000 rows
// Developer Intent: Compare the PDF export duration of a large workbook before and after applying chart localization via ChartGlobalizationSettings.
// Use Cases: Determine if custom axis labels affect PDF generation speed in high‑volume reports | Validate localization of chart elements while monitoring performance thresholds | Create automated regression tests for chart globalization impact on export time | Optimize large‑scale reporting pipelines that require localized charts
// AI Prompts: Generate a C# example that records export times for PDF and XLSX with and without ChartGlobalizationSettings using Aspose.Cells. | Explain how to interpret the timing results and suggest optimization techniques for chart localization performance. | Write a unit test that asserts the localized export time does not exceed a configurable percentage of the non‑localized export time.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPerformanceDemo
{
    // Custom chart globalization settings to simulate localization impact
    // Creates a workbook with 10,000 rows, adds a column chart, measures PDF export time, then applies a custom ChartGlobalizationSettings that localizes axis unit names, recalculates the chart, and measures the export time again. The two durations are printed for direct comparison.
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Example: change axis unit names to localized strings
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "百"; // Chinese for hundreds
                case DisplayUnitType.Thousands:
                    return "千"; // Chinese for thousands
                case DisplayUnitType.TenThousands:
                    return "万"; // Chinese for ten‑thousands
                default:
                    return base.GetAxisUnitName(type);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large dataset (e.g., 10,000 rows)
            int rowCount = 10000;
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            for (int i = 0; i < rowCount; i++)
            {
                cells[i + 1, 0].PutValue("Item " + (i + 1));
                cells[i + 1, 1].PutValue(i % 1000 + 1); // sample numeric data
            }

            // Add a column chart covering the data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 5, 25, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B" + (rowCount + 1), true);
            chart.NSeries.CategoryData = "A2:A" + (rowCount + 1);
            chart.Title.Text = "Large Data Chart";

            // Ensure chart layout is calculated before export
            chart.Calculate();

            // ------------------- Export without localization -------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Export to PDF (lifecycle: save)
            wb.Save("LargeWorkbook_NoLocalization.pdf", SaveFormat.Pdf);

            sw.Stop();
            Console.WriteLine($"Export without localization took: {sw.ElapsedMilliseconds} ms");

            // ------------------- Apply localization to charts -------------------
            // Set custom chart globalization settings
            wb.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new CustomChartGlobalizationSettings()
            };

            // Recalculate chart after applying localization (required for updated labels)
            chart.Calculate();

            // ------------------- Export with localization -------------------
            sw.Restart();

            // Export to PDF again
            wb.Save("LargeWorkbook_WithLocalization.pdf", SaveFormat.Pdf);

            sw.Stop();
            Console.WriteLine($"Export with localization took: {sw.ElapsedMilliseconds} ms");

            // Clean up
            wb.Dispose();
        }
    }
}
