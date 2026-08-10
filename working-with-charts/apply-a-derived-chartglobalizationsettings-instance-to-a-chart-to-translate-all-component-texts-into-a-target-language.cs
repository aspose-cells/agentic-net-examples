// Title: C# – Localize Excel Chart Texts with Aspose.Cells using ChartGlobalizationSettings
// Description: Shows how to build a workbook, insert a column chart, set French translations for the chart title, series name, legend actions and the "Other" label via SettableChartGlobalizationSettings, attach the settings to Workbook.Settings.GlobalizationSettings, and save the file so the chart renders with the localized strings.
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | SettableChartGlobalizationSettings | Excel chart localization | chart title translation | legend text globalization | multilingual Excel | programmatic chart localization | Aspose.Cells examples | Excel automation
// Common Searches: Aspose.Cells chart localization C# | SettableChartGlobalizationSettings example | translate Excel chart titles with Aspose | globalize chart legends in .NET | change chart text language programmatically | French chart labels Aspose.Cells | apply chart globalization settings to workbook | Excel chart multilingual support Aspose
// Developer Intent: Apply a derived ChartGlobalizationSettings instance to a chart so that every textual component (title, series, legend items, etc.) appears in a chosen target language.
// Use Cases: Produce French‑language sales dashboards where chart titles, series names, and legend actions are automatically translated. | Create a reusable SettableChartGlobalizationSettings object and apply it to multiple charts across worksheets for consistent multilingual branding. | Generate a single Excel report that contains charts in different languages by switching the workbook's GlobalizationSettings before saving each version.
// AI Prompts: Generate C# code that uses SettableChartGlobalizationSettings to localize chart texts to Spanish in an existing Aspose.Cells workbook. | Explain how to modify the ChartGlobalizationSettings of a saved workbook, reload it, and render the updated translations. | Provide a pattern for assigning distinct ChartGlobalizationSettings to separate worksheets so that each sheet displays charts in a different language.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalizationDemo
{
    // Shows how to build a workbook, insert a column chart, set French translations for the chart title, series name, legend actions and the "Other" label via SettableChartGlobalizationSettings, attach the settings to Workbook.Settings.GlobalizationSettings, and save the file so the chart renders with the localized strings.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");

            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Sales Overview";

            // Create a SettableChartGlobalizationSettings instance and set custom (French) texts
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetChartTitleName("Titre du graphique");   // Chart title
            chartSettings.SetSeriesName("Série");                   // Series name
            chartSettings.SetLegendIncreaseName("Augmenter");       // Legend increase
            chartSettings.SetLegendDecreaseName("Diminuer");        // Legend decrease
            chartSettings.SetLegendTotalName("Total");              // Legend total
            chartSettings.SetOtherName("Autre");                    // "Other" label

            // Apply the chart globalization settings to the workbook
            GlobalizationSettings globalization = new GlobalizationSettings
            {
                ChartSettings = chartSettings
            };
            workbook.Settings.GlobalizationSettings = globalization;

            // Save the workbook (the chart will display the localized texts)
            workbook.Save("ChartGlobalization_French.xlsx");
        }
    }
}
