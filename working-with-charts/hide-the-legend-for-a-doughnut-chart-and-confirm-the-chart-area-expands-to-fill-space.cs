using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class HideDoughnutChartLegend
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the doughnut chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(20);

                // Add a doughnut chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Doughnut, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the legend
                chart.ShowLegend = false;
                chart.Legend.IsOverLay = false; // optional, demonstrates usage

                // Save the workbook
                string outputPath = "DoughnutChart_NoLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideDoughnutChartLegend.Run();
        }
    }
}