// Title: Localize Aspose.Cells Chart Titles, Legends, and Axis Labels in C# (.NET) – Keep Data Values Intact
// Description: This example shows how to create a workbook with a column chart, then use Aspose.Cells SettableChartGlobalizationSettings to replace the chart title, axis titles, and legend entries with custom localized strings while leaving the numeric series unchanged. The workbook is saved as an Excel file ready for multilingual reporting.
// Keywords: Aspose.Cells chart localization C# | SettableChartGlobalizationSettings | localize chart title .NET | translate chart legend Aspose | axis label localization Excel | preserve data values Aspose.Cells | multilingual Excel charts | globalization settings C# | Aspose.Cells chart UI translation
// Common Searches: how to translate chart titles in Aspose.Cells C# | Aspose.Cells keep data values when localizing charts | set custom legend text for Excel chart using Aspose | globalize axis labels in Aspose.Cells .NET | example of chart localization with SettableChartGlobalizationSettings
// Developer Intent: Use Aspose.Cells globalization settings to apply localized text to chart UI elements (title, axis titles, legend) without modifying the underlying cell data.
// Use Cases: Generate Excel reports for international audiences where only chart captions need translation. | Automate creation of language‑specific workbooks that share the same data but display localized chart labels. | Build a single data source workbook and apply different SettableChartGlobalizationSettings for each target locale.
// AI Prompts: Write C# code with Aspose.Cells to localize chart titles, axis titles, and legend text while preserving numeric values. | List the SettableChartGlobalizationSettings properties that affect chart titles, axis labels, and legend entries. | Explain how to apply distinct localization strings to multiple charts in one workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalization
{
    // This example shows how to create a workbook with a column chart, then use Aspose.Cells SettableChartGlobalizationSettings to replace the chart title, axis titles, and legend entries with custom localized strings while leaving the numeric series unchanged. The workbook is saved as an Excel file ready for multilingual reporting.
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
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Sales Overview";       // Original (will be localized)

            // Configure axis titles (these will be localized)
            chart.CategoryAxis.Title.Text = "Quarter";
            chart.ValueAxis.Title.Text = "Revenue";

            // Configure legend (the legend text will be localized)
            chart.ShowLegend = true;

            // Create SettableChartGlobalizationSettings and set custom localized texts
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetChartTitleName("Localized Chart Title");
            chartSettings.SetAxisTitleName("Localized Axis Title");
            chartSettings.SetLegendIncreaseName("Localized Increase");
            chartSettings.SetLegendDecreaseName("Localized Decrease");
            chartSettings.SetLegendTotalName("Localized Total");

            // Assign the chart globalization settings to the workbook's GlobalizationSettings
            GlobalizationSettings globalization = new GlobalizationSettings
            {
                ChartSettings = chartSettings
            };
            workbook.Settings.GlobalizationSettings = globalization;

            // Save the workbook – data values remain unchanged, only titles/legends/axis labels are localized
            workbook.Save("LocalizedChart.xlsx");
        }
    }
}
