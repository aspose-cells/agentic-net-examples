using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChartRenderingPerformance
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Y values
        chart.NSeries.CategoryData = "A2:A4";      // X categories

        // Enable data labels and disable text wrapping for performance testing
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;
        dataLabels.IsTextWrapped = false; // disable wrapping

        // Measure the time taken to calculate (render) the chart
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        chart.Calculate(); // forces chart layout and rendering calculations
        stopwatch.Stop();

        Console.WriteLine($"Chart rendering calculation time: {stopwatch.ElapsedMilliseconds} ms");

        // Save the workbook with the chart
        workbook.Save("ChartPerformance.xlsx");
    }
}