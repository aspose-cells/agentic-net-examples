using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // ------------------------------------------------------------
        // 1. Create globalization settings for subtotal (total) label
        // ------------------------------------------------------------
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        // Change the label used for the Sum total (e.g., in subtotals)
        globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

        // ------------------------------------------------------------
        // 2. Create chart globalization settings for a pie chart
        // ------------------------------------------------------------
        SettableChartGlobalizationSettings chartGlobal = new SettableChartGlobalizationSettings();
        // Customize series name, chart title, legend total and "Other" label
        chartGlobal.SetSeriesName("Custom Series");
        chartGlobal.SetChartTitleName("Custom Pie Chart");
        chartGlobal.SetLegendTotalName("Custom Total");
        chartGlobal.SetOtherName("Other (Custom)");

        // Attach the chart settings to the main globalization object
        globalization.ChartSettings = chartGlobal;

        // ------------------------------------------------------------
        // 3. Apply the globalization settings to the workbook
        // ------------------------------------------------------------
        workbook.Settings.GlobalizationSettings = globalization;

        // ------------------------------------------------------------
        // 4. (Optional) Ensure there is a pie chart to demonstrate the effect
        // ------------------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        if (sheet.Charts.Count > 0)
        {
            Chart chart = sheet.Charts[0];
            chart.Type = ChartType.Pie;               // Force pie type for clarity
            chart.Title.Text = "Placeholder Title";   // Title will be replaced by globalization when default is used
        }

        // ------------------------------------------------------------
        // 5. Save the modified workbook
        // ------------------------------------------------------------
        workbook.Save("output.xlsx");
    }
}