// Title: Localize Aspose.Cells chart axis unit labels to German with a custom ChartGlobalizationSettings (C#)
// Description: Shows how to create a workbook, add sample data, build a column chart, set the value axis display unit to Thousands, enable the unit label, and apply a custom GermanChartGlobalizationSettings (overriding GetAxisUnitName) via GlobalizationSettings.ChartSettings so the axis label appears in German before saving the file.
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | German localization | chart axis unit label | DisplayUnitType | globalization settings | custom chart globalization | GermanChartGlobalizationSettings | Aspose.Cells chart localization
// Common Searches: Aspose.Cells German chart axis label | custom ChartGlobalizationSettings C# example | localize chart display unit label German | override GetAxisUnitName Aspose.Cells | set chart globalization settings .NET
// Developer Intent: Apply a custom ChartGlobalizationSettings object to translate chart axis unit names into German.
// Use Cases: Generate German‑language reports where chart axes show unit labels such as “Tausend” or “Millionen”. | Reuse GermanChartGlobalizationSettings across multiple workbooks for consistent chart localization. | Extend the class to support additional DisplayUnitType values or other languages. | Toggle back to the default globalization settings when German localization is no longer required.
// AI Prompts: Write C# code that defines a FrenchChartGlobalizationSettings class overriding GetAxisUnitName and applies it to a line chart in Aspose.Cells. | Explain how GlobalizationSettings.ChartSettings overrides axis unit names and describe the steps to revert to the default settings. | Provide a step‑by‑step method to programmatically verify that the German axis unit label appears correctly after the workbook is saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings that provide German names for axis units
    // Shows how to create a workbook, add sample data, build a column chart, set the value axis display unit to Thousands, enable the unit label, and apply a custom GermanChartGlobalizationSettings (overriding GetAxisUnitName) via GlobalizationSettings.ChartSettings so the axis label appears in German before saving the file.
    public class GermanChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Return German equivalents for common display unit types
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "Hundert";
                case DisplayUnitType.Thousands:
                    return "Tausend";
                case DisplayUnitType.Millions:
                    return "Millionen";
                case DisplayUnitType.Billions:
                    return "Milliarden";
                case DisplayUnitType.Percentage:
                    return "Prozent";
                default:
                    // Fallback to base implementation for other types
                    return base.GetAxisUnitName(type);
            }
        }
    }

    public class ApplyGermanChartGlobalization
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Kategorie");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Wert");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(250);
            sheet.Cells["B4"].PutValue(350);

            // Create a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the value axis to use a display unit and show its label
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Apply the custom German globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new GermanChartGlobalizationSettings()
            };

            // At this point, the axis unit label will be displayed in German
            Console.WriteLine("German axis unit label: " + chart.ValueAxis.DisplayUnitLabel.Text);

            // Save the workbook (adjust the path as needed)
            workbook.Save("GermanChartGlobalization.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ApplyGermanChartGlobalization.Run();
        }
    }
}
