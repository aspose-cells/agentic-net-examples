using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Sample Chart";

        // Attach custom globalization settings that log each method call
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new LoggingChartGlobalizationSettings()
        };

        // Force chart calculation so that labels are generated
        chart.Calculate();

        // Access legend labels to trigger label generation and observe logs
        Console.WriteLine("Legend Labels:");
        foreach (string label in chart.Legend.GetLegendLabels())
        {
            Console.WriteLine(label);
        }

        // Save the workbook
        workbook.Save("LoggingChartGlobalization.xlsx");
    }

    // Custom ChartGlobalizationSettings that logs when culture‑specific label methods are invoked
    class LoggingChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetLegendIncreaseName()
        {
            string result = base.GetLegendIncreaseName();
            Console.WriteLine("[Log] GetLegendIncreaseName called, returning: " + result);
            return result;
        }

        public override string GetLegendDecreaseName()
        {
            string result = base.GetLegendDecreaseName();
            Console.WriteLine("[Log] GetLegendDecreaseName called, returning: " + result);
            return result;
        }

        public override string GetLegendTotalName()
        {
            string result = base.GetLegendTotalName();
            Console.WriteLine("[Log] GetLegendTotalName called, returning: " + result);
            return result;
        }

        public override string GetOtherName()
        {
            string result = base.GetOtherName();
            Console.WriteLine("[Log] GetOtherName called, returning: " + result);
            return result;
        }

        public override string GetSeriesName()
        {
            string result = base.GetSeriesName();
            Console.WriteLine("[Log] GetSeriesName called, returning: " + result);
            return result;
        }

        public override string GetChartTitleName()
        {
            string result = base.GetChartTitleName();
            Console.WriteLine("[Log] GetChartTitleName called, returning: " + result);
            return result;
        }

        public override string GetAxisTitleName()
        {
            string result = base.GetAxisTitleName();
            Console.WriteLine("[Log] GetAxisTitleName called, returning: " + result);
            return result;
        }

        public override string GetAxisUnitName(DisplayUnitType type)
        {
            string result = base.GetAxisUnitName(type);
            Console.WriteLine("[Log] GetAxisUnitName called for " + type + ", returning: " + result);
            return result;
        }
    }
}