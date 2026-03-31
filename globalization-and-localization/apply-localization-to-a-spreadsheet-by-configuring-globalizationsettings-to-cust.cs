using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (Category and Value columns)
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        string[] categories = { "A", "B", "C", "D" };
        int[] values = { 10, 20, 30, 40 };
        for (int i = 0; i < categories.Length; i++)
        {
            cells[i + 1, 0].PutValue(categories[i]);   // Column A
            cells[i + 1, 1].PutValue(values[i]);      // Column B
        }

        // ------------------------------------------------------------
        // Configure globalization settings
        // ------------------------------------------------------------
        // 1. Create SettableGlobalizationSettings to customize subtotal label
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        // Custom total name for the Sum function (used by Subtotal)
        globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

        // 2. Create SettableChartGlobalizationSettings to customize chart labels
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetSeriesName("Custom Series");               // Series name in legend
        chartSettings.SetLegendTotalName("Custom Total");           // "Total" label in legend
        chartSettings.SetChartTitleName("Custom Pie Chart");        // Chart title
        // Assign chart settings to the main globalization object
        globalization.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // ------------------------------------------------------------
        // Create subtotal (adds a total row with the custom label)
        // ------------------------------------------------------------
        // Define the range that includes the data (including header row)
        CellArea dataArea = CellArea.CreateCellArea(0, 0, categories.Length, 1);
        // Subtotal on the first column (Category) using Sum on the Value column
        cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

        // ------------------------------------------------------------
        // Add a pie chart that uses the same data
        // ------------------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 10);
        Chart pieChart = sheet.Charts[chartIndex];
        // Add the values series; true indicates that the series is a column of values
        pieChart.NSeries.Add("B2:B5", true);
        // Set category (slice) labels
        pieChart.NSeries.CategoryData = "A2:A5";
        // The chart title will be overridden by the custom globalization setting
        pieChart.Title.Text = "Demo Pie Chart";

        // ------------------------------------------------------------
        // Save the workbook
        // ------------------------------------------------------------
        workbook.Save("LocalizedWorkbook.xlsx");
    }
}