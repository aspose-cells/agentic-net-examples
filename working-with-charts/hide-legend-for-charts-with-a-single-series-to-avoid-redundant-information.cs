using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class HideLegendForSingleSeriesChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a chart with a single series
                int chartIndex1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart1 = sheet.Charts[chartIndex1];
                chart1.NSeries.Add("B2:B4", true);
                chart1.NSeries.CategoryData = "A2:A4";

                // Add another chart with two series for demonstration
                int chartIndex2 = sheet.Charts.Add(ChartType.Column, 20, 0, 30, 5);
                Chart chart2 = sheet.Charts[chartIndex2];
                chart2.NSeries.Add("B2:B4", true);
                chart2.NSeries.Add("B2:B4", true); // second series (duplicate data for demo)
                chart2.NSeries.CategoryData = "A2:A4";

                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Hide legend if only one series, otherwise show it
                    chart.ShowLegend = chart.NSeries.Count != 1;
                }

                // Save the workbook
                string outputPath = "HideLegendSingleSeriesChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideLegendForSingleSeriesChart.Run();
        }
    }
}