// Title: Aspose.Cells C# – Verify Default English Labels with ChartGlobalizationSettings
// Description: Creates a workbook with a column chart, uses ChartGlobalizationSettings without custom values, and prints the built‑in English strings for series name, chart title, legend actions, axis title, other label, and axis unit (thousands). The sample checks that every string is non‑empty and saves the file, demonstrating the fallback localization behavior.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | default chart localization | English fallback labels | C# chart series name | GetChartTitleName | GetLegendIncreaseName | GetAxisUnitName | display unit thousands | chart globalization .NET
// Common Searches: Aspose.Cells default chart labels C# | How to get built‑in English text for chart elements in Aspose.Cells | ChartGlobalizationSettings default strings example | Verify chart localization when no globalization is set | Aspose.Cells fallback language for charts
// Developer Intent: Ensure that Aspose.Cells returns the built‑in English strings for all chart UI elements when no custom globalization configuration is applied.
// Use Cases: Display the default English series, title, legend, and axis labels for a newly generated chart. | Run a quick sanity check that all fallback localization strings are populated before exporting a workbook. | Create automated tests that confirm the presence of default chart texts in environments without localization settings.
// AI Prompts: Generate C# code that lists every default chart label provided by ChartGlobalizationSettings and asserts they are not null or empty. | Write a unit test for Aspose.Cells that creates a column chart, accesses ChartGlobalizationSettings, and verifies the English fallback strings. | Explain how ChartGlobalizationSettings supplies fallback English text for chart elements when no custom globalization is configured.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalizationDemo
{
    // Creates a workbook with a column chart, uses ChartGlobalizationSettings without custom values, and prints the built‑in English strings for series name, chart title, legend actions, axis title, other label, and axis unit (thousands). The sample checks that every string is non‑empty and saves the file, demonstrating the fallback localization behavior.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Do NOT assign any custom globalization settings.
            // Use the default ChartGlobalizationSettings to retrieve default English texts.
            ChartGlobalizationSettings defaultSettings = new ChartGlobalizationSettings();

            // Retrieve default texts
            string seriesName = defaultSettings.GetSeriesName();
            string chartTitleName = defaultSettings.GetChartTitleName();
            string legendIncreaseName = defaultSettings.GetLegendIncreaseName();
            string legendDecreaseName = defaultSettings.GetLegendDecreaseName();
            string legendTotalName = defaultSettings.GetLegendTotalName();
            string axisTitleName = defaultSettings.GetAxisTitleName();
            string otherName = defaultSettings.GetOtherName();
            string axisUnitName = defaultSettings.GetAxisUnitName(DisplayUnitType.Thousands);

            // Output the retrieved default English texts to the console
            Console.WriteLine("Default Series Name: " + seriesName);
            Console.WriteLine("Default Chart Title Name: " + chartTitleName);
            Console.WriteLine("Default Legend Increase Name: " + legendIncreaseName);
            Console.WriteLine("Default Legend Decrease Name: " + legendDecreaseName);
            Console.WriteLine("Default Legend Total Name: " + legendTotalName);
            Console.WriteLine("Default Axis Title Name: " + axisTitleName);
            Console.WriteLine("Default Other Name: " + otherName);
            Console.WriteLine("Default Axis Unit Name (Thousands): " + axisUnitName);

            // Simple verification: ensure none of the retrieved strings are null or empty
            bool allTextsPresent = !string.IsNullOrEmpty(seriesName) &&
                                   !string.IsNullOrEmpty(chartTitleName) &&
                                   !string.IsNullOrEmpty(legendIncreaseName) &&
                                   !string.IsNullOrEmpty(legendDecreaseName) &&
                                   !string.IsNullOrEmpty(legendTotalName) &&
                                   !string.IsNullOrEmpty(axisTitleName) &&
                                   !string.IsNullOrEmpty(otherName) &&
                                   !string.IsNullOrEmpty(axisUnitName);

            Console.WriteLine("All default English texts present: " + allTextsPresent);

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("DefaultChartLocalizationDemo.xlsx");
        }
    }
}
