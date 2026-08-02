using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CreateLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill cells A1 through A10 with sample numeric data
        for (int i = 0; i < 10; i++)
        {
            // Row index i (0‑based) corresponds to A{i+1}
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Add a line chart to the worksheet (positioned from rows 5‑20, columns 0‑5)
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart series (A1:A10)
        chart.NSeries.Add("=Sheet1!$A$1:$A$10", true);

        // Optional: set a chart title
        chart.Title.Text = "Sample Line Chart";

        // Save the workbook
        workbook.Save("LineChart.xlsx", SaveFormat.Xlsx);
    }
}