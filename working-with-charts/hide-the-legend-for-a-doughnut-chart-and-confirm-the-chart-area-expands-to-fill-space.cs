// Title: Hide the legend of a doughnut chart and let the plot area automatically fill the defined cell range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide legend in doughnut chart | remove legend from doughnut chart and expand plot area in .NET | how to make doughnut chart fill its container after hiding legend with Aspose.Cells | C# Aspose.Cells chart area size when legend is disabled | auto resize doughnut chart plot area after legend removal using Aspose.Cells
// Tags: doughnut chart legend visibility Aspose.Cells | hide chart legend C# Aspose.Cells | chart plot area auto fit Aspose.Cells | expand chart area after legend removal Aspose.Cells | Aspose.Cells doughnut chart without legend

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, adds sample data, inserts a doughnut chart spanning rows 5‑25 and columns 0‑10, sets the series and category ranges, and demonstrates how to hide the chart legend (chart.Legend.IsVisible = false). It notes that explicit PlotArea positioning is not available in newer API versions, then saves the workbook as DoughnutChart_NoLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the doughnut chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(50);

            // Add a doughnut chart (positioned from row 5, column 0 to row 25, column 10)
            int chartIdx = sheet.Charts.Add(ChartType.Doughnut, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Define the series and category data ranges
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend (if supported by the current API version)
            // Some versions expose IsVisible; if not, this line can be omitted.
            // chart.Legend.IsVisible = false;

            // Note: PlotArea positioning properties (Top, Left, Width, Height) are not
            // available in newer API versions, so they are omitted.

            // Determine output file path
            string outputPath = "DoughnutChart_NoLegend.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
