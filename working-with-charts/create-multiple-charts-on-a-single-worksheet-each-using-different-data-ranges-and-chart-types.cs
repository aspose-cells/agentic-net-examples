using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultipleChartsDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data for three different charts
                // Data for Column Chart (A1:B5)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Cat{i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                }

                // Data for Line Chart (A1:C5)
                sheet.Cells["C1"].PutValue("Series2");
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[$"C{i}"].PutValue(i * 15);
                }

                // Data for Pie Chart (A1:D5) – categories in A, values in D
                sheet.Cells["D1"].PutValue("PieValues");
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[$"D{i}"].PutValue(i * 5);
                }

                // -------------------------------------------------
                // Add first chart: Column Chart using range A1:B5
                // Position: rows 7-20, columns 0-5
                int colChartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 5);
                Chart colChart = sheet.Charts[colChartIndex];
                colChart.NSeries.Add("A1:B5", true);               // Data series
                colChart.NSeries.CategoryData = "A2:A5";          // Categories
                colChart.Title.Text = "Column Chart Example";

                // -------------------------------------------------
                // Add second chart: Line Chart using range A1:C5
                // Position: rows 22-35, columns 0-5
                int lineChartIndex = sheet.Charts.Add(ChartType.Line, 22, 0, 35, 5);
                Chart lineChart = sheet.Charts[lineChartIndex];
                lineChart.NSeries.Add("A1:C5", true);             // Two series (B and C columns)
                lineChart.NSeries.CategoryData = "A2:A5";        // Categories
                lineChart.Title.Text = "Line Chart Example";

                // -------------------------------------------------
                // Add third chart: Pie Chart using categories A2:A5 and values D2:D5
                // Position: rows 7-20, columns 7-12
                int pieChartIndex = sheet.Charts.Add(ChartType.Pie, 7, 7, 20, 12);
                Chart pieChart = sheet.Charts[pieChartIndex];
                // For pie chart, only one series is needed; set values and categories separately
                pieChart.NSeries.Add("D2:D5", false);             // Values
                pieChart.NSeries.CategoryData = "A2:A5";         // Categories
                pieChart.Title.Text = "Pie Chart Example";

                // -------------------------------------------------
                // Save the workbook to an XLSX file
                string outputPath = "MultipleChartsDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}