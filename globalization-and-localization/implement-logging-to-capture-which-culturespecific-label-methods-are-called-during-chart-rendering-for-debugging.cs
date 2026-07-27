using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDebugging
{
    // Custom globalization settings that log each method call
    public class LoggingChartGlobalizationSettings : SettableChartGlobalizationSettings
    {
        public override string GetLegendIncreaseName()
        {
            string result = base.GetLegendIncreaseName();
            Console.WriteLine("[LOG] GetLegendIncreaseName called, returning: " + result);
            return result;
        }

        public override string GetLegendDecreaseName()
        {
            string result = base.GetLegendDecreaseName();
            Console.WriteLine("[LOG] GetLegendDecreaseName called, returning: " + result);
            return result;
        }

        public override string GetLegendTotalName()
        {
            string result = base.GetLegendTotalName();
            Console.WriteLine("[LOG] GetLegendTotalName called, returning: " + result);
            return result;
        }

        public override string GetOtherName()
        {
            string result = base.GetOtherName();
            Console.WriteLine("[LOG] GetOtherName called, returning: " + result);
            return result;
        }

        public override string GetSeriesName()
        {
            string result = base.GetSeriesName();
            Console.WriteLine("[LOG] GetSeriesName called, returning: " + result);
            return result;
        }

        public override string GetChartTitleName()
        {
            string result = base.GetChartTitleName();
            Console.WriteLine("[LOG] GetChartTitleName called, returning: " + result);
            return result;
        }

        public override string GetAxisTitleName()
        {
            string result = base.GetAxisTitleName();
            Console.WriteLine("[LOG] GetAxisTitleName called, returning: " + result);
            return result;
        }

        public override string GetAxisUnitName(DisplayUnitType type)
        {
            string result = base.GetAxisUnitName(type);
            Console.WriteLine("[LOG] GetAxisUnitName called for " + type + ", returning: " + result);
            return result;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Quarterly Sales";

            // Apply custom logging globalization settings
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new LoggingChartGlobalizationSettings()
            };

            // Force chart calculation to trigger label generation
            chart.Calculate();

            // Access legend entries to ensure related methods are invoked
            Legend legend = chart.Legend;
            var legendLabels = legend.GetLegendLabels(); // modern API

            Console.WriteLine("Legend Labels:");
            foreach (string label in legendLabels)
            {
                Console.WriteLine(" - " + label);
            }

            // Save the workbook (lifecycle save)
            workbook.Save("ChartWithLogging.xlsx");
        }
    }
}