using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data that will be used as the chart source
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["A2"].PutValue("Cat1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Cat2");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cat3");
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the chart's data source range (rows 1‑4, columns A‑B)
        chart.SetChartDataRange("A1:B4", true);

        // Freeze the rows that contain the chart data source (first 4 rows)
        // FreezePanes(string cellName, int freezedRows, int freezedColumns)
        worksheet.FreezePanes("A5", 4, 0); // freezes rows 1‑4, no columns frozen

        // Save the workbook to a file
        workbook.Save("ChartWithFrozenSource.xlsx");
    }
}