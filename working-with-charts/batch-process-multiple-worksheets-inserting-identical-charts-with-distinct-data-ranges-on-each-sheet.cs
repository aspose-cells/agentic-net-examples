using Aspose.Cells;
using Aspose.Cells.Charts;

class BatchChartExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Number of worksheets to process
        int numberOfSheets = 3;

        // Loop to create worksheets, fill data, and add identical charts with sheet‑specific ranges
        for (int i = 0; i < numberOfSheets; i++)
        {
            Worksheet sheet;

            // First sheet already exists at index 0; subsequent sheets are added
            if (i == 0)
            {
                sheet = workbook.Worksheets[0];
                sheet.Name = $"Sheet{i + 1}";
            }
            else
            {
                sheet = workbook.Worksheets.Add($"Sheet{i + 1}");
            }

            // Populate sample data: column A = categories, column B = values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int row = 2; row <= 6; row++)
            {
                sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                // Distinct values per sheet to demonstrate different data ranges
                sheet.Cells[$"B{row}"].PutValue((row - 1) * (i + 1) * 10);
            }

            // Add a column chart to the worksheet.
            // Using the overload: Add(ChartType, string dataRange, bool isVertical, int topRow, int leftColumn, int rightRow, int bottomColumn)
            int chartIndex = sheet.Charts.Add(ChartType.Column, "A1:B6", true, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = $"Chart for Sheet {i + 1}";
        }

        // Save the workbook with all charts inserted
        workbook.Save("BatchCharts.xlsx");
    }
}