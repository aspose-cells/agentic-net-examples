// Title: Aspose.Cells for .NET – Localize Chart Titles, Legends & Axis Labels Only (Data Values Unchanged)
// Description: This C# example creates a workbook with a column chart, configures SettableChartGlobalizationSettings with Spanish UI texts (title, axis, legend), wraps it in a custom GlobalizationSettings class, assigns it to the workbook, and saves the file. Only the chart’s titles, legends, and axis labels are translated; the numeric data (120, 150, 180) stays the same.
// Keywords: Aspose.Cells | .NET chart localization | C# chart globalization | SettableChartGlobalizationSettings | custom GlobalizationSettings | chart title translation | legend localization | axis label localization | preserve chart data values | Spanish chart labels | Latin America reporting | GitHub Aspose.Cells example
// Common Searches: Aspose.Cells localize chart title only | How to translate chart legends in .NET | Keep chart data unchanged while localizing UI text | SettableChartGlobalizationSettings example | Custom GlobalizationSettings for Aspose.Cells charts
// Developer Intent: Apply Aspose.Cells globalization settings so that only chart UI elements (titles, legends, axis labels) are translated, leaving the underlying data untouched.
// Use Cases: Produce a sales column chart for a Spanish‑language report without altering the numeric sales figures. | Generate multilingual financial dashboards where each workbook’s chart labels are localized per locale while the data series remain identical. | Reuse the CustomGlobalizationSettings class to switch between language packs for chart UI across multiple workbooks in an automated pipeline.
// AI Prompts: Show how to set French labels using SettableChartGlobalizationSettings in Aspose.Cells. | Provide code that applies different CustomGlobalizationSettings to two charts in the same workbook. | Explain how to programmatically confirm that chart data values are unchanged after applying GlobalizationSettings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalizationDemo
{
    // Custom globalization settings that only changes chart titles, legends and axis labels.
    // This C# example creates a workbook with a column chart, configures SettableChartGlobalizationSettings with Spanish UI texts (title, axis, legend), wraps it in a custom GlobalizationSettings class, assigns it to the workbook, and saves the file. Only the chart’s titles, legends, and axis labels are translated; the numeric data (120, 150, 180) stays the same.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Constructor receives the prepared SettableChartGlobalizationSettings instance.
        public CustomGlobalizationSettings(SettableChartGlobalizationSettings chartSettings)
        {
            // Assign the chart settings to the base GlobalizationSettings.
            this.ChartSettings = chartSettings;
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create ----------
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ---------- Localization ----------
            // Create an instance of SettableChartGlobalizationSettings.
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();

            // Set custom localized texts for titles, legends and axis labels.
            chartSettings.SetChartTitleName("Ventas Mensuales");          // Chart title
            chartSettings.SetAxisTitleName("Meses");                     // Axis title (both axes use same text here)
            chartSettings.SetLegendIncreaseName("Aumento");              // Legend increase label
            chartSettings.SetLegendDecreaseName("Disminución");          // Legend decrease label
            chartSettings.SetLegendTotalName("Total");                   // Legend total label
            chartSettings.SetOtherName("Otros");                         // "Other" label (if used)

            // Apply the custom globalization settings to the workbook.
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings(chartSettings);

            // The data values (e.g., 120, 150, 180) remain unchanged; only UI texts are localized.

            // ---------- Save ----------
            // Save the workbook to an XLSX file.
            workbook.Save("LocalizedChart.xlsx");
        }
    }
}
