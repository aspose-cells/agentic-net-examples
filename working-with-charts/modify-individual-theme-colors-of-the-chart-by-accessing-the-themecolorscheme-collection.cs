using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsThemeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");

                // Populate sample data for the chart
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Category {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);   // Sample value for Series 1
                    sheet.Cells[$"C{i}"].PutValue(i * 15);   // Sample value for Series 2
                }

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:C6", true);          // Data range
                chart.NSeries.CategoryData = "A2:A6";      // Category (X) axis
                chart.Title.Text = "Sample Column Chart";

                // Define output file path
                string outputPath = Path.Combine(Environment.CurrentDirectory, "DemoChart.xlsx");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}