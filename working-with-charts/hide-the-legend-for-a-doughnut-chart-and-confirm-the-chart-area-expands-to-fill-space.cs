// Title: Hide the legend of a doughnut chart and make the chart area fill its container with Aspose.Cells for .NET
// AI Prompts: Generate a new workbook, add sample data, insert a doughnut chart, turn off its legend, and verify that the plot area occupies the entire chart range using Aspose.Cells for .NET. | Take an existing doughnut chart created with Aspose.Cells, disable the legend display, and programmatically expand the chart’s plot area to use all available space.
// Common Searches: Aspose.Cells C# hide legend doughnut chart and expand plot area | programmatically remove legend from Excel doughnut chart using Aspose.Cells | auto‑resize chart area after disabling legend in Aspose.Cells .NET | how to make doughnut chart fill its range when legend is hidden in C# | Aspose.Cells chart layout without legend for Excel export
// Tags: Aspose.Cells suppress chart legend C# | doughnut chart plot area expansion Aspose.Cells | chart layout adjustment after legend removal .NET | Aspose.Cells auto resize chart area Excel | C# remove legend doughnut chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, populates it with sample data, adds a doughnut chart, assigns the data series, disables the legend by setting ShowLegend to false, and saves the file as DoughnutChart.xlsx, resulting in a chart that expands to fill the defined area.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data for the doughnut chart
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(50);

            // Add a doughnut chart (positioned from row 5, column 0 to row 25, column 10)
            int chartIndex = ws.Charts.Add(ChartType.Doughnut, 5, 0, 25, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend
            chart.ShowLegend = false;

            // Save the workbook
            string outputPath = "DoughnutChart.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
