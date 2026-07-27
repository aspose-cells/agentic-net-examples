using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class TenSheetsWithCharts
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
            // Create a new workbook (no template file needed)
            Workbook workbook = new Workbook();

            // Loop to create 10 worksheets, each with its own data and chart
            for (int i = 1; i <= 10; i++)
            {
                // Add a new worksheet with a unique name
                string sheetName = $"Sheet{i}";
                Worksheet sheet = workbook.Worksheets.Add(sheetName);

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Data rows (A‑E categories)
                for (int row = 2; row <= 6; row++)
                {
                    char category = (char)('A' + row - 2);
                    sheet.Cells[$"A{row}"].PutValue(category.ToString());
                    sheet.Cells[$"B{row}"].PutValue(i * 10 + (row - 2));
                }

                // Add a column chart to the current worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Bind the chart to the data range of this worksheet
                chart.NSeries.Add($"{sheetName}!B2:B6", true);
                chart.NSeries.CategoryData = $"{sheetName}!A2:A6";

                // Optional: give the chart a title
                chart.Title.Text = $"Chart {i}";
            }

            // Save the workbook containing all worksheets and charts
            string outputPath = "TenSheetsWithCharts.xlsx";

            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}