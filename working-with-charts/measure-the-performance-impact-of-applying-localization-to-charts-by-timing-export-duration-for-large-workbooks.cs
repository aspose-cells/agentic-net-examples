// Title: Benchmark Chart Localization Export Performance with Aspose.Cells for .NET
// Description: C# sample that builds a 2,000‑row by 10‑column workbook with a column chart, measures the XLSX save time with Stopwatch, then applies a custom ChartGlobalizationSettings (e.g., Chinese axis unit names), recalculates the chart, and records the export duration again to show the overhead of chart localization.
// Keywords: Aspose.Cells | C# chart localization | performance benchmark | export time measurement | large workbook | ChartGlobalizationSettings | Stopwatch timing | XLSX save speed | globalization impact | axis unit translation
// Common Searches: Aspose.Cells chart localization performance test | measure export time for large workbook with charts | benchmark Aspose.Cells XLSX save with custom globalization | how much does chart globalization slow down export | C# timing Aspose.Cells chart export
// Developer Intent: Find out how custom chart globalization settings influence the time required to save a large workbook using Aspose.Cells.
// Use Cases: Compare export durations of a 2,000‑row workbook with and without a custom ChartGlobalizationSettings implementation. | Validate that invoking chart.Calculate() after changing globalization ensures correct rendering before saving. | Create automated performance tests for different SaveFormat values (XLSX, PDF, PNG) while applying chart localization.
// AI Prompts: Write C# code that logs export times for XLSX, PDF, and PNG formats while applying a custom ChartGlobalizationSettings to charts in a large workbook. | Explain the internal steps Aspose.Cells follows when processing chart globalization during workbook save and suggest ways to reduce latency. | Generate an NUnit test that asserts the export‑time difference between localized and non‑localized charts stays below a configurable threshold.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalizationPerformance
{
    // Custom globalization settings for charts (e.g., change axis unit names)
    // C# sample that builds a 2,000‑row by 10‑column workbook with a column chart, measures the XLSX save time with Stopwatch, then applies a custom ChartGlobalizationSettings (e.g., Chinese axis unit names), recalculates the chart, and records the export duration again to show the overhead of chart localization.
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Example: translate unit names to another language
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
        // Creates a large workbook with sample data and a column chart
        static Workbook CreateLargeWorkbook(int rows, int columns)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    // Simple numeric data; first column as categories
                    if (c == 0)
                        cells[r, c].PutValue($"Item {r + 1}");
                    else
                        cells[r, c].PutValue((r + 1) * (c + 1));
                }
            }

            // Add a column chart covering the data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, rows + 2, 0, rows + 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set series data (all numeric columns)
            string firstDataCol = CellsHelper.ColumnIndexToName(1);
            string lastDataCol = CellsHelper.ColumnIndexToName(columns - 1);
            chart.NSeries.Add($"{firstDataCol}1:{lastDataCol}{rows}", true);
            chart.NSeries.CategoryData = $"A1:A{rows}";
            chart.Title.Text = "Sample Large Chart";

            // Ensure chart layout is calculated before export
            chart.Calculate();

            return wb;
        }

        // Measures export time for a given workbook and save format
        static long MeasureExportTime(Workbook wb, string filePath, SaveFormat format)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Save the workbook (lifecycle rule: save)
            wb.Save(filePath, format);

            sw.Stop();
            return sw.ElapsedMilliseconds;
        }

        static void Main()
        {
            try
            {
                // Parameters for a "large" workbook
                const int rowCount = 2000;   // adjust as needed for testing
                const int colCount = 10;

                // 1. Export without localization
                Workbook wbNoLoc = CreateLargeWorkbook(rowCount, colCount);
                string outNoLoc = "LargeWorkbook_NoLocalization.xlsx";

                // Ensure we can write to the output path
                if (File.Exists(outNoLoc))
                    File.Delete(outNoLoc);

                long timeNoLoc = MeasureExportTime(wbNoLoc, outNoLoc, SaveFormat.Xlsx);
                Console.WriteLine($"Export without localization: {timeNoLoc} ms");

                // 2. Apply chart localization settings
                Workbook wbLoc = CreateLargeWorkbook(rowCount, colCount);
                wbLoc.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new CustomChartGlobalizationSettings()
                };

                // Force chart to re-calculate after globalization change
                foreach (Chart chart in wbLoc.Worksheets[0].Charts)
                {
                    chart.Calculate();
                }

                string outLoc = "LargeWorkbook_WithLocalization.xlsx";

                if (File.Exists(outLoc))
                    File.Delete(outLoc);

                long timeLoc = MeasureExportTime(wbLoc, outLoc, SaveFormat.Xlsx);
                Console.WriteLine($"Export with localization: {timeLoc} ms");

                // Simple performance comparison
                long diff = timeLoc - timeNoLoc;
                Console.WriteLine(diff >= 0
                    ? $"Localization added {diff} ms overhead."
                    : $"Localization reduced export time by {-diff} ms.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
