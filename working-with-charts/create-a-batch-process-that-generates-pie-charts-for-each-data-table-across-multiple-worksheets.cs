using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchPieCharts
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (you can also load an existing file if needed)
            Workbook workbook = new Workbook();

            // Example: add three worksheets with sample data tables
            for (int i = 0; i < 3; i++)
            {
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                // Populate sample data: Column A = Category, Column B = Value
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int row = 2; row <= 6; row++)
                {
                    sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                    sheet.Cells[$"B{row}"].PutValue(row * 10 + i * 5); // varied values per sheet
                }

                // Add a pie chart to the worksheet
                // Position the chart from row 8, column 0 to row 25, column 7
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 8, 0, 25, 7);
                Chart pieChart = sheet.Charts[chartIndex];

                // Define the data range for the series (values) and categories
                // Values are in B2:B6, categories in A2:A6
                pieChart.NSeries.Add("B2:B6", true);
                pieChart.NSeries.CategoryData = "A2:A6";

                // Optional: set chart title
                pieChart.Title.Text = $"Pie Chart for {sheet.Name}";
            }

            // Save the workbook to a file
            workbook.Save("BatchPieCharts.xlsx");
        }
    }
}