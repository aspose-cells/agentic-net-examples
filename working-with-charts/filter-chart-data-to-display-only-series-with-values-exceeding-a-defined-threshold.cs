// Title: Filter chart series by numeric threshold with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart, and hide any series whose values never exceed a defined threshold. The example uses the Series.IsFiltered property and shows how to count filtered series before saving the file.
// Keywords: Aspose.Cells chart filtering C# | hide chart series Aspose.Cells | Series.IsFiltered property | threshold based chart series | .NET Excel chart example | filter NSeries Aspose | Excel chart automation C# | dynamic chart series visibility
// Common Searches: Aspose.Cells hide chart series below a threshold | C# filter column chart series by value | How to use Series.IsFiltered in Aspose.Cells | Remove low‑value series from Excel chart programmatically | Aspose.Cells chart series conditional display
// Developer Intent: Show only those chart series that contain at least one data point greater than a specified numeric limit.
// Use Cases: Sales dashboards that automatically omit products with sales below target. | KPI reports where only metrics surpassing risk thresholds appear in charts. | Monthly performance sheets that exclude insignificant data series without manual editing.
// AI Prompts: Generate C# code using Aspose.Cells to hide chart series whose all values are below a given threshold. | Explain the role of the IsFiltered property for chart series and how to retrieve the filtered series count after applying a threshold. | Provide a pattern for making the threshold user‑configurable and re‑applying the filter without recreating the chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartFiltering
{
    // Demonstrates how to create a workbook, add a column chart, and hide any series whose values never exceed a defined threshold. The example uses the Series.IsFiltered property and shows how to count filtered series before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: three series (B, C, D) with numeric values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(5);
                sheet.Cells["C3"].PutValue(15);
                sheet.Cells["C4"].PutValue(25);

                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(40);
                sheet.Cells["D3"].PutValue(50);
                sheet.Cells["D4"].PutValue(60);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Add each series to the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries.Add("D2:D4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Define the threshold: only series containing a value > 25 will be shown
                double threshold = 25.0;

                // Iterate through each series and filter out those that do not exceed the threshold
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    // Get the range string that holds the series values (e.g., "B2:B4")
                    string rangeStr = series.Values;

                    // Create a range object to access individual cells
                    AsposeRange range = sheet.Cells.CreateRange(rangeStr);

                    bool exceedsThreshold = false;
                    foreach (Cell cell in range)
                    {
                        // Ensure the cell contains a numeric value before comparison
                        if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > threshold)
                        {
                            exceedsThreshold = true;
                            break;
                        }
                    }

                    // If no value exceeds the threshold, hide the series
                    if (!exceedsThreshold)
                    {
                        series.IsFiltered = true;
                    }
                }

                // Optional: display count of filtered series
                Console.WriteLine("Filtered series count: " + chart.FilteredNSeries.Count);

                // Save the workbook
                workbook.Save("ChartFilteredByThreshold.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
