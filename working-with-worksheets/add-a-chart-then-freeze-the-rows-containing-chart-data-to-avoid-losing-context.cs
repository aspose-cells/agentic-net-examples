using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart (rows 1‑6, columns A‑B)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 5; i++)
        {
            sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
            sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
        }

        // Add a column chart positioned below the data (rows 7‑20, columns A‑I)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart's data source to the range we just filled
        chart.NSeries.Add("=Sheet1!$A$1:$B$6", true);

        // Freeze the rows and columns that contain the chart data
        // Freeze first 6 rows and first 2 columns (A and B)
        sheet.FreezePanes(6, 2, 6, 2);

        // Save the workbook
        workbook.Save("ChartWithFrozenRows.xlsx", SaveFormat.Xlsx);
    }
}