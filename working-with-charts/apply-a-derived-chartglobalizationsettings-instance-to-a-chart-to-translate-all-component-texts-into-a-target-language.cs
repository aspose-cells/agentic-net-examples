// Title: Localize Aspose.Cells Chart Texts with SettableChartGlobalizationSettings (C#)
// Description: Demonstrates how to create a SettableChartGlobalizationSettings object, define French labels for chart title, series, legend actions and the "Other" term, embed it in SettableGlobalizationSettings, assign it to a Workbook, and generate a column chart that automatically uses the localized strings before saving the file.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | SettableChartGlobalizationSettings | chart localization .NET | Excel chart translation | multilingual Excel | French chart labels | globalization settings | C# Aspose.Cells example
// Common Searches: Aspose.Cells how to localize chart titles | Set custom chart globalization settings in C# | Translate chart legends with Aspose.Cells | Apply French labels to Excel charts using Aspose | Globalize chart text in .NET workbook
// Developer Intent: Apply a custom ChartGlobalizationSettings object so that every chart component displays text in the chosen language.
// Use Cases: Generate French‑language sales reports where chart titles, series names, and legend actions appear in French without manual text replacement. | Create a reusable Excel template that automatically localizes chart UI strings for multiple languages via workbook globalization settings. | Produce multilingual workbooks where each chart inherits language‑specific terms such as "Total" and "Other" from a single configuration.
// AI Prompts: Show how to switch the example to Spanish localization for all chart component texts. | Provide code to read the applied ChartGlobalizationSettings from an existing workbook. | Explain how to assign different ChartGlobalizationSettings to separate charts in the same worksheet. | Suggest a way to load localization strings from a resource file and apply them to ChartGlobalizationSettings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalizationDemo
{
    // Demonstrates how to create a SettableChartGlobalizationSettings object, define French labels for chart title, series, legend actions and the "Other" term, embed it in SettableGlobalizationSettings, assign it to a Workbook, and generate a column chart that automatically uses the localized strings before saving the file.
    class Program
    {
        static void Main()
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

            // ------------------------------------------------------------
            // Create custom chart globalization settings (e.g., French)
            // ------------------------------------------------------------
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();

            // Set custom localized texts for various chart components
            chartSettings.SetChartTitleName("Titre du graphique");      // Chart title
            chartSettings.SetSeriesName("Série");                     // Series name
            chartSettings.SetLegendIncreaseName("Augmenter");         // Legend increase
            chartSettings.SetLegendDecreaseName("Diminuer");          // Legend decrease
            chartSettings.SetLegendTotalName("Total");                // Legend total
            chartSettings.SetOtherName("Autre");                      // "Other" label

            // Wrap the chart settings inside a SettableGlobalizationSettings instance
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();
            globalSettings.ChartSettings = chartSettings;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // ------------------------------------------------------------
            // Create a chart that will use the above globalization settings
            // ------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series data and category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set a placeholder title; the actual displayed title will use the localized name
            chart.Title.Text = "Placeholder Title";

            // Save the workbook
            workbook.Save("ChartWithCustomGlobalization.xlsx");
        }
    }
}
