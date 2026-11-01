// Title: Validate that ChartPoint.IsInSecondaryPlot is false for primary series points in an Aspose.Cells column chart (C#)
// AI Prompts: Generate a C# console program using Aspose.Cells that creates a column chart, loops through each point in the first series, and asserts the IsInSecondaryPlot property returns false. | Write code to programmatically confirm that no ChartPoint in a primary plot is marked as belonging to a secondary axis, throwing an exception if the check fails.
// Common Searches: how to ensure ChartPoint.IsInSecondaryPlot returns false for column chart series in Aspose.Cells | Aspose.Cells C# verify that chart points are not assigned to secondary axis | sample code to check secondary plot flag on chart points using Aspose.Cells | validate primary series points in a column chart with Aspose.Cells .NET
// Tags: primary plot point verification Aspose.Cells | column chart series iteration C# | detect secondary axis points Aspose.Cells | ChartPoint property validation Aspose.Cells | Aspose.Cells chart point status check

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, builds a column chart on the primary plot, accesses the first series, iterates over each ChartPoint, checks the IsInSecondaryPlot property, throws an exception if any point reports true, and saves the workbook.
class ChartPointValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart (primary plot)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data source for the chart series
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Access the first series of the chart
            var series = chart.NSeries[0];

            // Validate each point is not in the secondary plot
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                bool isInSecondary = point.IsInSecondaryPlot;
                Console.WriteLine($"Point Index: {i}, IsInSecondaryPlot: {isInSecondary}");

                if (isInSecondary)
                {
                    throw new InvalidOperationException(
                        $"ChartPoint at index {i} incorrectly reports being in the secondary plot.");
                }
            }

            // Save the workbook
            string outputPath = "ChartPointValidation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
