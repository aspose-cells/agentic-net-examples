using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill cells A1 through A10 with sample numeric data
        for (int i = 0; i < 10; i++)
        {
            // Row index i, column index 0 corresponds to column A
            sheet.Cells[i, 0].PutValue(i + 1); // values 1,2,...,10
        }

        // Add a line chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (vertical series)
        chart.SetChartDataRange("A1:A10", true);

        // Optional: set a chart title
        chart.Title.Text = "Line Chart from A1:A10";

        // Save the workbook to a file
        workbook.Save("LineChart.xlsx", SaveFormat.Xlsx);
    }
}