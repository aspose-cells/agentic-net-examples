// Title: Verify default English chart labels in Aspose.Cells when ChartGlobalizationSettings are not set (C#)
// Description: This example creates a workbook with a column chart, uses a default ChartGlobalizationSettings instance (no culture assigned), extracts the built‑in English strings for series name, chart title, legend items, axis title and unit, prints them, and saves the file. It demonstrates that Aspose.Cells falls back to English localization when no custom globalization is applied.
// Keywords: Aspose.Cells default chart localization | ChartGlobalizationSettings English strings | C# Aspose chart default language | retrieve default chart labels Aspose.Cells | default axis unit name Aspose.Cells
// Common Searches: Aspose.Cells default chart language | How to get built‑in English chart labels in Aspose.Cells | ChartGlobalizationSettings default values .NET | Check chart localization without custom settings Aspose | Default axis unit name Aspose.Cells C#
// Developer Intent: Confirm that a chart displays the built‑in English labels when no ChartGlobalizationSettings are configured.
// Use Cases: Log default chart strings to ensure correct language fallback before applying custom localization. | Create unit tests that validate the English defaults returned by ChartGlobalizationSettings. | Generate reports that rely on the standard English labels when multilingual support is not required.
// AI Prompts: Generate a C# unit test that asserts each default string from ChartGlobalizationSettings matches the expected English text. | Show how to switch ChartGlobalizationSettings to French and compare the resulting strings with the English defaults. | Explain how to programmatically verify that a saved workbook’s chart uses default English labels when no globalization settings are applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook with a column chart, uses a default ChartGlobalizationSettings instance (no culture assigned), extracts the built‑in English strings for series name, chart title, legend items, axis title and unit, prints them, and saves the file. It demonstrates that Aspose.Cells falls back to English localization when no custom globalization is applied.
class VerifyDefaultChartLocalization
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Instantiate default ChartGlobalizationSettings (no custom localization assigned)
        ChartGlobalizationSettings defaultSettings = new ChartGlobalizationSettings();

        // Retrieve default English texts from the settings
        string seriesName = defaultSettings.GetSeriesName();
        string chartTitleName = defaultSettings.GetChartTitleName();
        string legendIncrease = defaultSettings.GetLegendIncreaseName();
        string legendDecrease = defaultSettings.GetLegendDecreaseName();
        string legendTotal = defaultSettings.GetLegendTotalName();
        string axisTitle = defaultSettings.GetAxisTitleName();
        string otherName = defaultSettings.GetOtherName();
        string axisUnit = defaultSettings.GetAxisUnitName(DisplayUnitType.Thousands);

        // Output the retrieved default texts to verify they are English
        Console.WriteLine("Default Series Name: " + seriesName);
        Console.WriteLine("Default Chart Title Name: " + chartTitleName);
        Console.WriteLine("Default Legend Increase Name: " + legendIncrease);
        Console.WriteLine("Default Legend Decrease Name: " + legendDecrease);
        Console.WriteLine("Default Legend Total Name: " + legendTotal);
        Console.WriteLine("Default Axis Title Name: " + axisTitle);
        Console.WriteLine("Default Other Name: " + otherName);
        Console.WriteLine("Default Axis Unit (Thousands): " + axisUnit);

        // Save the workbook (no custom globalization settings applied)
        workbook.Save("DefaultChartLocalization.xlsx");
    }
}
