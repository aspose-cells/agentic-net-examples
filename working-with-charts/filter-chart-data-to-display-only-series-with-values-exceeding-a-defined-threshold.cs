// Title: Aspose.Cells C# – Hide Chart Series Below a Numeric Threshold
// Description: Creates a workbook, adds a column chart with two series, calculates each series' maximum value, and uses the IsFiltered flag to hide any series whose maximum does not exceed a defined threshold. The example also shows how to read the filtered series count and save the file.
// Keywords: Aspose.Cells chart filtering | C# hide chart series | IsFiltered property | chart series threshold | FilteredNSeries Aspose.Cells | compute max value range | dynamic chart visibility
// Common Searches: Aspose.Cells hide chart series below threshold | C# filter column chart series by max value | How to use IsFiltered in Aspose.Cells chart | Get count of filtered series Aspose.Cells | Calculate max value of a range for chart filtering
// Developer Intent: Programmatically hide chart series that do not meet a numeric threshold.
// Use Cases: Automatically exclude low‑sales product lines from a sales chart. | Generate dashboards that only display KPI‑meeting series. | Create reports that suppress insignificant data series before saving.
// AI Prompts: Generate code to filter chart series in Aspose.Cells based on the average value instead of the maximum. | Show how to change the threshold at runtime and refresh chart visibility without rebuilding the series collection. | Provide a method to unfilter all series and then apply a new threshold in an existing Aspose.Cells chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartFiltering
{
    // Creates a workbook, adds a column chart with two series, calculates each series' maximum value, and uses the IsFiltered flag to hide any series whose maximum does not exceed a defined threshold. The example also shows how to read the filtered series count and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Series 1 values
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2 values
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Add both series to the chart
                chart.NSeries.Add("B2:B4", true); // Series1
                chart.NSeries.Add("C2:C4", true); // Series2
                chart.NSeries.CategoryData = "A2:A4";

                // Define the threshold – only series whose maximum value exceeds this will be shown
                double threshold = 25.0;

                // Iterate through each series, evaluate its data, and filter if needed
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series ser = chart.NSeries[i];

                    // The series data range is stored in the Values property (e.g., "B2:B4")
                    string dataRange = ser.Values;

                    // Compute the maximum value within the range
                    double maxVal = GetMaxValueFromRange(sheet, dataRange);

                    // Hide series whose maximum does not exceed the threshold
                    ser.IsFiltered = maxVal <= threshold;
                }

                // Optionally, you can access the collection of filtered (hidden) series
                SeriesCollection filtered = chart.FilteredNSeries;
                Console.WriteLine($"Number of filtered (hidden) series: {filtered.Count}");

                // Save the workbook
                string outputPath = "ChartFilteredByThreshold.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to compute the maximum numeric value in a given range string (e.g., "B2:B4")
        private static double GetMaxValueFromRange(Worksheet sheet, string range)
        {
            // Create a Range object from the address string
            AsposeRange cellsRange = sheet.Cells.CreateRange(range);

            double max = double.MinValue;

            foreach (Cell cell in cellsRange)
            {
                if (cell.Value is double d)
                {
                    if (d > max) max = d;
                }
                else if (cell.Value is int i)
                {
                    double dVal = i;
                    if (dVal > max) max = dVal;
                }
            }

            // If no numeric values were found, treat max as 0
            return max == double.MinValue ? 0.0 : max;
        }
    }
}
