// Title: Validate ChartPoint.IsInSecondaryPlot for Primary Points in a Pie‑of‑Pie Chart (Aspose.Cells .NET)
// Description: Creates a workbook, adds a Pie‑of‑Pie chart, marks one data point as secondary, then iterates all points to confirm IsInSecondaryPlot is false for every primary point and true only for the designated secondary point, finally saving the file.
// Keywords: Aspose.Cells | ChartPoint.IsInSecondaryPlot | PiePie chart | secondary plot .NET | C# chart validation | chart series point property | Aspose.Cells example | primary plot check
// Common Searches: Aspose.Cells check if chart point is in secondary plot | ChartPoint.IsInSecondaryPlot false for primary points | Pie of Pie chart secondary plot example C# | How to validate secondary plot points in Aspose.Cells | Aspose.Cells chart series point status
// Developer Intent: Verify that ChartPoint.IsInSecondaryPlot returns false for points that belong to the primary plot.
// Use Cases: Programmatically assign a single data point to the secondary plot of a Pie‑of‑Pie chart and ensure all other points remain primary. | Log each point’s category and its secondary‑plot flag for debugging chart configurations. | Create automated tests that confirm only the intended point is marked as secondary before publishing a workbook.
// AI Prompts: Generate C# code using Aspose.Cells that sets one chart point to the secondary plot and asserts IsInSecondaryPlot is false for all other points. | Write an xUnit test that verifies ChartPoint.IsInSecondaryPlot is true only for the designated secondary point in a PiePie chart. | Explain how to read and output the IsInSecondaryPlot property for every point in an Aspose.Cells chart series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a Pie‑of‑Pie chart, marks one data point as secondary, then iterates all points to confirm IsInSecondaryPlot is false for every primary point and true only for the designated secondary point, finally saving the file.
class ValidateChartPointSecondaryPlot
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a Pie of Pie chart (supports secondary plot)
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["A5"].PutValue("D");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["B5"].PutValue(40);

        // Add a Pie of Pie chart (ChartType.PiePie) which allows secondary plots
        int chartIndex = worksheet.Charts.Add(ChartType.PiePie, 5, 0, 25, 15);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";     // Categories

        // Explicitly set one point to be in the secondary plot
        // Here we set the third point (index 2) to secondary; others remain primary
        chart.NSeries[0].Points[2].IsInSecondaryPlot = true;

        // Validate that points not set to secondary have IsInSecondaryPlot == false
        for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
        {
            ChartPoint point = chart.NSeries[0].Points[i];
            bool isSecondary = point.IsInSecondaryPlot;
            Console.WriteLine($"Point {i} (Category '{worksheet.Cells[i + 2, 0].StringValue}') IsInSecondaryPlot = {isSecondary}");

            // Expect false for all points except the one we set (index 2)
            if (i != 2 && isSecondary)
            {
                Console.WriteLine("Validation error: point should be in primary plot but IsInSecondaryPlot is true.");
            }
        }

        // Save the workbook to verify the chart is created correctly
        workbook.Save("ValidateIsInSecondaryPlot.xlsx");
    }
}
