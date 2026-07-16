using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: continuous numeric X values and corresponding Y values
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 11; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i - 1);               // X = 1,2,...,10
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 5);        // Y = 5,10,...,50
        }

        // Add a scatter chart (ideal for numeric X axis) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range: Y values and X values
        chart.NSeries.Add("B2:B11", true);
        chart.NSeries[0].XValues = "A2:A11";

        // Retrieve the X (category) axis and change its type to continuous numeric scaling.
        // AutomaticScale treats the axis as a value axis suitable for continuous data.
        chart.CategoryAxis.CategoryType = CategoryType.AutomaticScale;

        // Save the workbook with the modified chart
        workbook.Save("ChartWithNumericXAxis.xlsx", SaveFormat.Xlsx);
    }
}