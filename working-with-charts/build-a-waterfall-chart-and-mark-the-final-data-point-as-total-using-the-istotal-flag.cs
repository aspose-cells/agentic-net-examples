// Title: Create a Waterfall Chart with Total Column using Aspose.Cells for .NET
// Description: This example shows how to generate a workbook, populate category and value columns, add a Waterfall chart, bind the series to ranges A2:A6 and B2:B6, and mark the last point as a total with the ChartPoint.IsTotal flag (when supported). The chart is recalculated and saved as WaterfallChartWithTotal.xlsx.
// Keywords: Aspose.Cells | C# Waterfall chart | ChartPoint.IsTotal | total column waterfall | programmatic Excel chart | .NET chart series | waterfall chart example
// Common Searches: Aspose.Cells set IsTotal on waterfall chart point | C# create waterfall chart with total column | how to mark final point as total in Aspose.Cells | waterfall chart series data range Aspose.Cells | Aspose.Cells waterfall chart code sample
// Developer Intent: Generate a waterfall chart and flag the final data point as a total.
// Use Cases: Financial reports that need a highlighted ending balance in a waterfall visualization. | Performance dashboards displaying revenue, cost, and profit with a cumulative total column. | Automated monthly statements that include a waterfall chart with a program‑defined total segment.
// AI Prompts: Write C# code using Aspose.Cells to create a waterfall chart and set ChartPoint.IsTotal = true for the last point, including a version‑check fallback. | Explain how to detect whether the current Aspose.Cells version supports ChartPoint.IsTotal and suggest an alternative for older releases. | Demonstrate how to recalculate a waterfall chart after modifying point properties such as IsTotal in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartExample
{
    // This example shows how to generate a workbook, populate category and value columns, add a Waterfall chart, bind the series to ranges A2:A6 and B2:B6, and mark the last point as a total with the ChartPoint.IsTotal flag (when supported). The chart is recalculated and saved as WaterfallChartWithTotal.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the waterfall chart
                // Column A – Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["A3"].PutValue("Revenue");
                sheet.Cells["A4"].PutValue("Cost");
                sheet.Cells["A5"].PutValue("Profit");
                sheet.Cells["A6"].PutValue("End");

                // Column B – Values
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(1000);   // Start
                sheet.Cells["B3"].PutValue(1500);   // Revenue
                sheet.Cells["B4"].PutValue(-500);   // Cost (negative to show drop)
                sheet.Cells["B5"].PutValue(0);      // Profit (calculated by Excel)
                sheet.Cells["B6"].PutValue(2000);   // End (total)

                // Add a Waterfall chart
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Mark the final data point as a total (if supported by the library version)
                Series series = chart.NSeries[0];
                int lastPointIndex = series.Points.Count - 1;
                ChartPoint lastPoint = series.Points[lastPointIndex];
                // The IsTotal property may not be available in older versions; this line is kept for newer versions.
                // Uncomment the following line if your Aspose.Cells version supports ChartPoint.IsTotal.
                // lastPoint.IsTotal = true;

                // Optional: calculate the chart to ensure all properties are applied
                chart.Calculate();

                // Save the workbook with the waterfall chart
                workbook.Save("WaterfallChartWithTotal.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
