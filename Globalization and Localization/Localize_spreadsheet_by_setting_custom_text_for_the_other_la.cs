using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing workbook (XLSX format)
        Workbook workbook = new Workbook("input.xlsx");

        // Create chart globalization settings and set custom text for the "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Custom Other");

        // Create a SettableGlobalizationSettings instance and assign the chart settings
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}