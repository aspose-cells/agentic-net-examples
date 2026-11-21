// Title: How to use ChartPoint.IsInSecondaryPlot to identify secondary axis points in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# check if a chart point is plotted on the secondary axis | Iterate chart series and detect secondary plot points with ChartPoint.IsInSecondaryPlot | Identify secondary axis data points in an Aspose.Cells column chart using ChartPoint.IsInSecondaryPlot | Read CategoryData and ValueData for secondary series points in Aspose.Cells .NET | Set a series to the secondary axis and verify points in Aspose.Cells C#
// Tags: ChartPoint.IsInSecondaryPlot detection | secondary axis point identification Aspose.Cells | iterate chart points Aspose.Cells C# | column chart secondary series Aspose.Cells | reflection set secondary axis Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook with primary and secondary data series, configures the secondary series to plot on the secondary axis using reflection, then iterates through each ChartPoint, reads its category and value via reflection, and uses ChartPoint.IsInSecondaryPlot to determine whether the point belongs to the secondary plot before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate data for primary and secondary series
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Primary");
            ws.Cells["C1"].PutValue("Secondary");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);
            ws.Cells["C2"].PutValue(100);
            ws.Cells["C3"].PutValue(200);
            ws.Cells["C4"].PutValue(300);

            // Add a column chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Primary series (plotted on primary axis)
            int primarySeriesIndex = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[primarySeriesIndex].Name = "Primary Series";

            // Secondary series (plotted on secondary axis)
            int secondarySeriesIndex = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[secondarySeriesIndex].Name = "Secondary Series";

            // Set the secondary series to plot on the secondary axis.
            // Use reflection to handle possible API differences.
            var secondarySeries = chart.NSeries[secondarySeriesIndex];
            var seriesType = secondarySeries.GetType();

            var isPlotProp = seriesType.GetProperty("IsPlotOnSecondAxis");
            if (isPlotProp != null && isPlotProp.CanWrite)
            {
                isPlotProp.SetValue(secondarySeries, true);
            }
            else
            {
                var altProp = seriesType.GetProperty("PlotOnSecondAxis");
                if (altProp != null && altProp.CanWrite)
                {
                    altProp.SetValue(secondarySeries, true);
                }
            }

            // Iterate through all points and display whether they belong to the secondary plot
            foreach (Series series in chart.NSeries)
            {
                foreach (ChartPoint point in series.Points)
                {
                    bool isInSecondary = point.IsInSecondaryPlot;

                    // Retrieve category and value via reflection to stay compatible with different versions
                    var catProp = point.GetType().GetProperty("CategoryData");
                    var valProp = point.GetType().GetProperty("ValueData");
                    string category = catProp?.GetValue(point)?.ToString() ?? "N/A";
                    string value = valProp?.GetValue(point)?.ToString() ?? "N/A";

                    Console.WriteLine($"Series: {series.Name}, Category: {category}, Value: {value}, IsInSecondaryPlot: {isInSecondary}");
                }
            }

            // Save the workbook
            string outputPath = "ChartPointSecondaryPlot.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
