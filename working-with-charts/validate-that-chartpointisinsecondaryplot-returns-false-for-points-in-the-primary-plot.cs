// Title: C# Aspose.Cells example: confirm ChartPoint.IsInSecondaryPlot is false for primary plot points in a column chart
// Common Searches: how to check ChartPoint.IsInSecondaryPlot for primary series points using Aspose.Cells in C# | Aspose.Cells C# code to validate that chart points are not in secondary plot | verify primary plot chart point property IsInSecondaryPlot with Aspose.Cells .NET | C# iterate column chart series points and test IsInSecondaryPlot flag Aspose.Cells | Aspose.Cells chart point secondary plot detection example in C#
// Tags: chartpoint isinsecondayplot false verification | aspose.cells column chart primary series check | c# chart series point property validation | aspose.cells chart point secondary plot detection | verify chart point plot membership aspnet cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds data, inserts a column chart with a primary series, then loops through each ChartPoint in the series to ensure the IsInSecondaryPlot property is false, throwing an exception if any point is incorrectly flagged as belonging to a secondary plot.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data for the chart (primary plot)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Item 1");
        sheet.Cells["A3"].PutValue("Item 2");
        sheet.Cells["A4"].PutValue("Item 3");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the series (primary plot)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series
        Series series = chart.NSeries[0];

        // Validate that each point belongs to the primary plot
        for (int i = 0; i < series.Points.Count; i++)
        {
            ChartPoint point = series.Points[i];
            bool isInSecondary = point.IsInSecondaryPlot;

            // Expected: false for primary plot points
            Console.WriteLine($"Point {i + 1} IsInSecondaryPlot = {isInSecondary}");
            if (isInSecondary)
            {
                throw new Exception($"Validation failed: Point {i + 1} is incorrectly marked as secondary plot.");
            }
        }

        Console.WriteLine("All points correctly reported as primary plot (IsInSecondaryPlot = false).");
    }
}
