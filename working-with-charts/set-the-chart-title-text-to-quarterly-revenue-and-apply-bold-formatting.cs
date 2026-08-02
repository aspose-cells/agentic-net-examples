using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetChartTitleBoldDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Quarter");
                sheet.Cells["B1"].PutValue("Revenue");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["B2"].PutValue(150000);
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["B3"].PutValue(200000);
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B4"].PutValue(180000);
                sheet.Cells["A5"].PutValue("Q4");
                sheet.Cells["B5"].PutValue(220000);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Set chart title text and make it bold
                chart.Title.Text = "Quarterly Revenue";
                chart.Title.IsVisible = true;
                chart.Title.Font.IsBold = true;

                // Save the workbook
                string outputPath = "SetChartTitleBoldDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
                throw;
            }
        }
    }
}