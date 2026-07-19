// Title: Localize Aspose.Cells Chart Axis Unit Labels to German with SettableChartGlobalizationSettings (C#)
// Description: Demonstrates how to create a workbook, add a column chart, enable a display unit on the value axis, define German names for hundreds, thousands, millions and percentages using SettableChartGlobalizationSettings, apply the settings through workbook.Settings.GlobalizationSettings, and save the file so the chart shows German axis unit labels.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | SettableChartGlobalizationSettings | German localization | axis unit label | display unit | C# | Excel chart localization | .NET | Aspose.Cells chart example
// Common Searches: Aspose.Cells set chart axis label language | German axis unit label Aspose.Cells C# | How to use SettableChartGlobalizationSettings in .NET | Localize Excel chart units with Aspose.Cells | Change chart display unit text to German
// Developer Intent: Apply a custom ChartGlobalizationSettings object so that a chart’s axis unit labels appear in German.
// Use Cases: Generate a column chart that displays values in hundreds and shows the German label "Hundert". | Define German unit names for hundreds, thousands, millions, and percentages and apply them workbook‑wide. | Create an Excel file where all chart axis unit labels are localized for German‑speaking users. | Programmatically verify the localized unit label via the chart’s DisplayUnitLabel property.
// AI Prompts: Show how to localize chart axis unit labels to French using SettableChartGlobalizationSettings in Aspose.Cells for .NET. | Provide a step‑by‑step guide to apply different ChartGlobalizationSettings to multiple charts in the same workbook. | Explain how to read, modify, and refresh the DisplayUnitLabel text after changing globalization settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalization
{
    // Custom class that derives from SettableChartGlobalizationSettings
    // and sets German names for axis unit labels.
    // Demonstrates how to create a workbook, add a column chart, enable a display unit on the value axis, define German names for hundreds, thousands, millions and percentages using SettableChartGlobalizationSettings, apply the settings through workbook.Settings.GlobalizationSettings, and save the file so the chart shows German axis unit labels.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(250);
            sheet.Cells["B4"].PutValue(370);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Umsatz";

            // Configure the chart to display a unit label
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Hundreds;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Create SettableChartGlobalizationSettings and set German unit names
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetAxisUnitName(DisplayUnitType.Hundreds, "Hundert");
            chartSettings.SetAxisUnitName(DisplayUnitType.Thousands, "Tausend");
            chartSettings.SetAxisUnitName(DisplayUnitType.Millions, "Millionen");
            chartSettings.SetAxisUnitName(DisplayUnitType.Percentage, "Prozent");

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = chartSettings
            };

            // Verify the custom label (optional console output)
            Console.WriteLine("Axis unit label (German): " + chart.ValueAxis.DisplayUnitLabel.Text);

            // Save the workbook
            workbook.Save("ChartWithGermanAxisLabels.xlsx");
        }
    }
}
