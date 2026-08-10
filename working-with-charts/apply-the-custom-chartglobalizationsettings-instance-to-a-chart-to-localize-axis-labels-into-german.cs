// Title: Localize Aspose.Cells Chart Axis Unit Labels to German with a Custom ChartGlobalizationSettings (C#)
// Description: This example shows how to create a GermanChartGlobalizationSettings class that inherits SettableChartGlobalizationSettings, assigns German names to display units (Hundert, Tausend, Millionen, Prozent), and applies the settings to a workbook via workbook.Settings.GlobalizationSettings. The chart’s value axis then displays German unit labels when the DisplayUnit is set to Hundreds or Thousands.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | SettableChartGlobalizationSettings | German localization | axis unit label | DisplayUnitType | C# | .NET | chart localization | workbook globalization | GermanChartGlobalizationSettings
// Common Searches: Aspose.Cells German chart axis labels | How to set chart globalization settings in C# | Localize chart display unit text to German | SettableChartGlobalizationSettings example | Change chart value axis unit label language Aspose.Cells
// Developer Intent: Apply a custom ChartGlobalizationSettings object to a workbook so that chart axis unit labels are rendered in German.
// Use Cases: Provide German‑language financial charts for European audiences. | Reuse a GermanChartGlobalizationSettings class across multiple workbooks. | Dynamically switch display units while keeping German unit labels consistent.
// AI Prompts: Generate C# code that defines a SettableChartGlobalizationSettings subclass for French axis unit names and applies it to an Aspose.Cells chart. | Show how to read the current axis unit label after changing the DisplayUnit in Aspose.Cells. | Explain how to configure workbook‑wide chart globalization for several languages using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalization
{
    // Custom class that configures German axis unit names using SettableChartGlobalizationSettings
    // This example shows how to create a GermanChartGlobalizationSettings class that inherits SettableChartGlobalizationSettings, assigns German names to display units (Hundert, Tausend, Millionen, Prozent), and applies the settings to a workbook via workbook.Settings.GlobalizationSettings. The chart’s value axis then displays German unit labels when the DisplayUnit is set to Hundreds or Thousands.
    class GermanChartGlobalizationSettings : SettableChartGlobalizationSettings
    {
        public GermanChartGlobalizationSettings()
        {
            // Set German names for common display unit types
            SetAxisUnitName(DisplayUnitType.Hundreds, "Hundert");
            SetAxisUnitName(DisplayUnitType.Thousands, "Tausend");
            SetAxisUnitName(DisplayUnitType.Millions, "Millionen");
            SetAxisUnitName(DisplayUnitType.Percentage, "Prozent");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Kategorie");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Wert");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(2500);
            sheet.Cells["B4"].PutValue(3700);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the value axis to use a display unit (Hundreds) and show the unit label
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Hundreds;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Apply German globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new GermanChartGlobalizationSettings()
            };

            // At this point, the axis unit label will be displayed in German ("Hundert")
            Console.WriteLine("Axis unit label (German): " + chart.ValueAxis.DisplayUnitLabel.Text);

            // Change the display unit to Thousands to demonstrate another German label
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            Console.WriteLine("Updated axis unit label (German): " + chart.ValueAxis.DisplayUnitLabel.Text);

            // Save the workbook
            workbook.Save("GermanLocalizedChart.xlsx");
        }
    }
}
