using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    public class SetCategoryDataDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 8; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                }

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the series values (vertical orientation)
                chart.NSeries.Add("B2:B8", true);

                // Define the category labels
                chart.NSeries.CategoryData = "A2:A8";

                // Save the workbook
                string outputPath = "CategoryDataDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetCategoryDataDemo.Run();
        }
    }
}