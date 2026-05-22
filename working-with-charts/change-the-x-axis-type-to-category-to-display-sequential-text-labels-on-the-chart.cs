using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetCategoryAxisTypeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1)); // sequential text labels
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                }

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and the category (X) axis
                chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

                // Change the X axis (category axis) type to CategoryScale
                chart.CategoryAxis.CategoryType = CategoryType.CategoryScale;

                // Save the workbook
                string outputPath = "SetCategoryAxisTypeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            SetCategoryAxisTypeDemo.Run();
        }
    }
}