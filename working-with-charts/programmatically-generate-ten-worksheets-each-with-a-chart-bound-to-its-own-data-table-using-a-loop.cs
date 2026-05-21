using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTenSheetsWithCharts
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (first worksheet is created by default)
            Workbook workbook = new Workbook();

            // Loop to create 10 worksheets, each with its own data and chart
            for (int i = 0; i < 10; i++)
            {
                // Use the default first worksheet for i == 0, otherwise add a new one
                Worksheet sheet;
                if (i == 0)
                {
                    sheet = workbook.Worksheets[0];
                    sheet.Name = $"Sheet{i + 1}";
                }
                else
                {
                    sheet = workbook.Worksheets.Add($"Sheet{i + 1}");
                }

                // Populate a simple data table: Category (A) and Value (B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Add 4 rows of sample data; values are varied per worksheet
                for (int row = 2; row <= 5; row++)
                {
                    sheet.Cells[row - 1, 0].PutValue($"Item {row - 1}");
                    // Example value: (worksheet index + 1) * row * 10
                    sheet.Cells[row - 1, 1].PutValue((i + 1) * row * 10);
                }

                // Add a column chart to the worksheet
                // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Bind the chart to the data range created above
                // Series values (Y) are in column B, categories (X) are in column A
                chart.NSeries.Add($"{sheet.Name}!B2:B5", true);
                chart.NSeries.CategoryData = $"{sheet.Name}!A2:A5";

                // Optional: set a title for each chart
                chart.Title.Text = $"Chart {i + 1}";
            }

            // Save the workbook containing all worksheets and charts
            workbook.Save("TenSheetsWithCharts.xlsx");
        }
    }
}