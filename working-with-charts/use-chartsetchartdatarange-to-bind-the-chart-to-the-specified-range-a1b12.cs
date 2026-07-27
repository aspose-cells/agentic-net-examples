using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in the range A1:B12
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 12; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            worksheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Bind the chart to the data range A1:B12 (vertical = true means plot by column)
        chart.SetChartDataRange("A1:B12", true);

        // Save the workbook
        workbook.Save("ChartWithRange.xlsx");
    }
}