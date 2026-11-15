// Title: How to reference a range on another worksheet as the data source for a column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column chart series to a range on a different worksheet | reference DataSheet!A1:B5 for chart series in Aspose.Cells .NET | create a chart on one sheet using data from another sheet with Aspose.Cells | cross‑sheet chart data source example Aspose.Cells for .NET | Aspose.Cells NSeries.Add external worksheet range C#
// Tags: Aspose.Cells set chart series range across worksheets | Aspose.Cells column chart data from another sheet | Aspose.Cells NSeries.Add external range | Aspose.Cells cross‑sheet chart linking | Aspose.Cells chart data source reference

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook with a 'DataSheet' containing sample values, adds a 'ChartSheet', inserts a column chart, and uses NSeries.Add to set the chart's data series to the range A1:B5 on 'DataSheet', demonstrating cross‑sheet chart data linking in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Worksheet that will contain the chart
            var chartSheet = workbook.Worksheets[0];
            chartSheet.Name = "ChartSheet";

            // Add a new worksheet for source data
            var dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            dataSheet.Name = "DataSheet";

            // Fill sample data in DataSheet (range A1:B5)
            double[,] sampleData = new double[,] { { 1, 10 }, { 2, 20 }, { 3, 30 }, { 4, 40 }, { 5, 50 } };
            for (int row = 0; row < 5; row++)
            {
                dataSheet.Cells[row, 0].PutValue(sampleData[row, 0]); // X values
                dataSheet.Cells[row, 1].PutValue(sampleData[row, 1]); // Y values
            }

            // Add a column chart to ChartSheet
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            var chart = chartSheet.Charts[chartIndex];
            chart.Title.Text = "Cross‑Sheet Data Chart";

            // Reference the data range on DataSheet for the series
            chart.NSeries.Add("DataSheet!A1:B5", true);

            // Save the workbook
            workbook.Save("CrossSheetChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
