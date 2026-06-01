using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ColumnChartDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(45);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the series
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Set custom data labels for each point
                Series series = chart.NSeries[0];
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    point.DataLabels.Text = point.YValue.ToString();
                    point.DataLabels.ShowValue = true;
                }

                // Save the workbook
                string outputPath = "ColumnChartWithDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ColumnChartDataLabels.Run();
        }
    }
}