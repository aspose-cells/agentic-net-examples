using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Do NOT assign any custom globalization settings.
            // Use the default ChartGlobalizationSettings to retrieve default English texts.
            ChartGlobalizationSettings defaultSettings = new ChartGlobalizationSettings();

            // Retrieve default texts
            string seriesName = defaultSettings.GetSeriesName();
            string chartTitleName = defaultSettings.GetChartTitleName();
            string legendIncreaseName = defaultSettings.GetLegendIncreaseName();
            string legendDecreaseName = defaultSettings.GetLegendDecreaseName();
            string legendTotalName = defaultSettings.GetLegendTotalName();
            string axisTitleName = defaultSettings.GetAxisTitleName();
            string otherName = defaultSettings.GetOtherName();
            string axisUnitName = defaultSettings.GetAxisUnitName(DisplayUnitType.Thousands);

            // Output the retrieved default English texts to the console
            Console.WriteLine("Default Series Name: " + seriesName);
            Console.WriteLine("Default Chart Title Name: " + chartTitleName);
            Console.WriteLine("Default Legend Increase Name: " + legendIncreaseName);
            Console.WriteLine("Default Legend Decrease Name: " + legendDecreaseName);
            Console.WriteLine("Default Legend Total Name: " + legendTotalName);
            Console.WriteLine("Default Axis Title Name: " + axisTitleName);
            Console.WriteLine("Default Other Name: " + otherName);
            Console.WriteLine("Default Axis Unit Name (Thousands): " + axisUnitName);

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("DefaultChartLocalizationDemo.xlsx");
        }
    }
}