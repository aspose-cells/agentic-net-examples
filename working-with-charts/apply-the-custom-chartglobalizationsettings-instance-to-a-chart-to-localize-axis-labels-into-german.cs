using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

public class GermanChartGlobalizationDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate sample data for the chart
        ws.Cells["A1"].PutValue("Kategorie");
        ws.Cells["A2"].PutValue("Q1");
        ws.Cells["A3"].PutValue("Q2");
        ws.Cells["A4"].PutValue("Q3");
        ws.Cells["B1"].PutValue("Wert");
        ws.Cells["B2"].PutValue(100);
        ws.Cells["B3"].PutValue(200);
        ws.Cells["B4"].PutValue(300);

        // Add a column chart and bind data
        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set a display unit and enable its label
        chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
        chart.ValueAxis.IsDisplayUnitLabelShown = true;

        // Create SettableChartGlobalizationSettings and define German unit names
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetAxisUnitName(DisplayUnitType.Hundreds, "Hundert");
        chartSettings.SetAxisUnitName(DisplayUnitType.Thousands, "Tausend");
        chartSettings.SetAxisUnitName(DisplayUnitType.Millions, "Millionen");
        chartSettings.SetAxisUnitName(DisplayUnitType.Percentage, "Prozent");

        // Apply the globalization settings to the workbook
        wb.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = chartSettings
        };

        // Verify the localized unit label
        Console.WriteLine("Localized display unit label: " + chart.ValueAxis.DisplayUnitLabel.Text);

        // Change the unit to demonstrate another German label
        chart.ValueAxis.DisplayUnit = DisplayUnitType.Hundreds;
        Console.WriteLine("Updated localized display unit label: " + chart.ValueAxis.DisplayUnitLabel.Text);

        // Save the workbook
        wb.Save("GermanChartGlobalizationDemo.xlsx");
    }
}