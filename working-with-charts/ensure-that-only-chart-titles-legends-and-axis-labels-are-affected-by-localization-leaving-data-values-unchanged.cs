// Title: Localize Chart Titles, Legends, and Axis Labels with Aspose.Cells for .NET while Keeping Data Values Intact
// Description: Shows how to employ Aspose.Cells' SettableChartGlobalizationSettings in C# to translate a column chart’s title, axis titles, and legend entries (e.g., to Spanish) without modifying the numeric series. The sample builds a workbook, inserts sales data, configures localization, creates the chart, and saves the workbook.
// Keywords: Aspose.Cells chart localization | SettableChartGlobalizationSettings | C# Excel chart translation | preserve chart data values | .NET multilingual Excel | translate chart title and legend | axis label localization Aspose | Excel column chart example | globalize chart UI text | Aspose.Cells GitHub example
// Common Searches: Aspose.Cells localize chart title without changing data | SettableChartGlobalizationSettings C# example | Translate Excel chart axis labels with Aspose.Cells | How to keep chart data values when localizing legends | Multilingual chart creation using Aspose.Cells .NET
// Developer Intent: Apply localization exclusively to chart textual elements (title, axis titles, legend) while leaving the underlying data untouched.
// Use Cases: Produce regional sales reports where only the chart captions appear in the target language, preserving original numeric formats. | Generate a single Excel workbook that can be re‑localized for different markets by swapping chart UI text without re‑exporting data. | Integrate automated chart translation into a reporting pipeline, ensuring data integrity while adapting visual labels for global audiences.
// AI Prompts: Write C# code using Aspose.Cells to translate chart titles, axis titles, and legend entries to French while keeping the data series unchanged. | Explain how to load localization strings from a .resx file and apply them to SettableChartGlobalizationSettings for an existing chart. | Provide a testing approach to verify that only chart text is affected by globalization settings and that cell values remain the same.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalizationDemo
{
    // Shows how to employ Aspose.Cells' SettableChartGlobalizationSettings in C# to translate a column chart’s title, axis titles, and legend entries (e.g., to Spanish) without modifying the numeric series. The sample builds a workbook, inserts sales data, configures localization, creates the chart, and saves the workbook.
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
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Configure localization for chart elements
                var chartLocalization = new SettableChartGlobalizationSettings();
                chartLocalization.SetChartTitleName("Ventas Mensuales");   // Chart title
                chartLocalization.SetAxisTitleName("Meses");              // Axis title (both axes)
                chartLocalization.SetLegendIncreaseName("Incremento");    // Legend increase
                chartLocalization.SetLegendDecreaseName("Decremento");    // Legend decrease
                chartLocalization.SetLegendTotalName("Total");            // Legend total

                // Assign the localized settings to the workbook
                workbook.Settings.GlobalizationSettings.ChartSettings = chartLocalization;

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Bind data series (values) and categories (labels)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Define output file path
                string outputPath = "LocalizedChartDemo.xlsx";

                // Ensure the directory exists (in case a relative path is used)
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
