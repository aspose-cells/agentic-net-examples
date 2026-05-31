using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a column chart to the worksheet (rows 5-20, columns 1-8)
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 8);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the chart (adjust as needed)
        chart.SetChartDataRange("A1:B5", true);

        // Optional: set a title for the chart
        chart.Title.Text = "Sample Column Chart";

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}