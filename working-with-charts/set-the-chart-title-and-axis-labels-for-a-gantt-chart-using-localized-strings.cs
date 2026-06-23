using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetChartLocalizationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for a chart
            ws.Cells["A1"].PutValue("Task");
            ws.Cells["B1"].PutValue("Start");
            ws.Cells["C1"].PutValue("Duration");

            ws.Cells["A2"].PutValue("Task 1");
            ws.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            ws.Cells["C2"].PutValue(5);

            ws.Cells["A3"].PutValue("Task 2");
            ws.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
            ws.Cells["C3"].PutValue(7);

            // Create globalization settings with localized strings
            SettableChartGlobalizationSettings chartGlobals = new SettableChartGlobalizationSettings();
            chartGlobals.SetChartTitleName("项目进度");      // Localized chart title
            chartGlobals.SetAxisTitleName("时间轴");        // Localized axis title

            // Apply the globalization settings to the workbook
            wb.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = chartGlobals
            };

            // Add a supported chart type (e.g., Column) to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the chart
            // Category (tasks) from A2:A3, values from B2:C3
            chart.NSeries.Add("{B2:C3}", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Apply localized titles
            chart.Title.Text = chartGlobals.GetChartTitleName();
            chart.Title.IsVisible = true;

            chart.CategoryAxis.Title.Text = chartGlobals.GetAxisTitleName();
            chart.CategoryAxis.Title.IsVisible = true;

            chart.ValueAxis.Title.Text = chartGlobals.GetAxisTitleName();
            chart.ValueAxis.Title.IsVisible = true;

            // Save the workbook
            string outputPath = "GanttLocalized.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}