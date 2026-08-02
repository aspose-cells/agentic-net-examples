using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchPieCharts
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define number of worksheets to process
            int sheetCount = 3;

            // Populate each worksheet with a simple data table
            for (int i = 0; i < sheetCount; i++)
            {
                // Add a new worksheet (first sheet already exists)
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Sample data rows (5 rows per sheet)
                for (int row = 2; row <= 6; row++)
                {
                    sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                    sheet.Cells[$"B{row}"].PutValue((row - 1) * 10 + i * 5); // Vary values per sheet
                }

                // Add a pie chart for this worksheet
                // Chart positioned from row 8, column 0 to row 20, column 7
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 8, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (values column)
                // The second parameter 'true' indicates that the first row contains category names
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Optional: set chart title
                chart.Title.Text = $"Pie Chart - {sheet.Name}";
            }

            // Save the workbook to a file
            workbook.Save("BatchPieCharts.xlsx");
        }
    }
}