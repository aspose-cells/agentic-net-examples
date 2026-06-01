using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalizationDemo
{
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add series data (values) and category data (labels)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ------------------------------------------------------------
            // Configure localization for chart elements (title, axis, legend)
            // ------------------------------------------------------------

            // Create a SettableChartGlobalizationSettings instance
            SettableChartGlobalizationSettings chartGlobals = new SettableChartGlobalizationSettings();

            // Set localized texts for the elements we want to affect
            chartGlobals.SetChartTitleName("Localized Chart Title");
            chartGlobals.SetAxisTitleName("Localized Axis Title");
            chartGlobals.SetSeriesName("Localized Series");
            chartGlobals.SetLegendIncreaseName("Localized Increase");
            chartGlobals.SetLegendDecreaseName("Localized Decrease");
            chartGlobals.SetLegendTotalName("Localized Total");

            // Assign the globalization settings to the workbook
            GlobalizationSettings globalization = new GlobalizationSettings
            {
                ChartSettings = chartGlobals
            };
            workbook.Settings.GlobalizationSettings = globalization;

            // Apply the localized texts to the chart
            chart.Title.Text = chartGlobals.GetChartTitleName();                     // Chart title
            chart.ValueAxis.Title.Text = chartGlobals.GetAxisTitleName();           // Value axis title
            chart.CategoryAxis.Title.Text = chartGlobals.GetAxisTitleName();        // Category axis title

            // Apply localized series name (affects legend entry for the series)
            if (chart.NSeries.Count > 0)
            {
                chart.NSeries[0].Name = chartGlobals.GetSeriesName();
            }

            // Optionally customize legend entries (increase/decrease/total) if needed
            // Aspose.Cells does not expose direct properties for these built‑in legend items,
            // but they will be used automatically when the chart type supports them.

            // ------------------------------------------------------------
            // Save the workbook (data values remain unchanged)
            // ------------------------------------------------------------
            workbook.Save("LocalizedChartDemo.xlsx");
        }
    }
}