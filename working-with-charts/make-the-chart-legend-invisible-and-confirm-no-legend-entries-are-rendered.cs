using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Hide the legend
        chart.ShowLegend = false;
        Console.WriteLine("Chart.ShowLegend after hiding: " + chart.ShowLegend);

        // Verify that legend entries are not rendered (legend is hidden)
        // The entries still exist in the collection, but they are not displayed because ShowLegend is false
        int entryCount = chart.Legend.LegendEntries?.Count ?? 0;
        Console.WriteLine("Legend entries count (still present in model): " + entryCount);

        // Save the workbook
        workbook.Save("ChartWithoutLegend.xlsx");
    }
}