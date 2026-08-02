// Title: C# – Hide Aspose.Cells Chart Series Below a Value Threshold
// Description: Creates a workbook with three data series, adds a column chart, then evaluates each series' maximum cell value. Series whose maximum does not exceed a defined threshold are hidden by setting the IsFiltered property. The example also shows how to read the filtered series count and save the workbook.
// Keywords: Aspose.Cells chart filtering C# | hide chart series Aspose.Cells | IsFiltered property | filter series by threshold | column chart series max value | .NET Excel chart automation | dynamic chart series visibility
// Common Searches: Aspose.Cells hide chart series below threshold | C# filter Excel chart series by value | How to use IsFiltered in Aspose.Cells | Remove low‑value series from Aspose.Cells chart | Count filtered series Aspose.Cells
// Developer Intent: Programmatically hide chart series whose data never exceeds a specified numeric threshold.
// Use Cases: Display only product lines with sales above a target in a sales dashboard. | Automatically exclude low‑risk assets from a financial performance chart. | Create a KPI report that shows only metrics surpassing a defined benchmark.
// AI Prompts: Generate C# code that uses Aspose.Cells to hide chart series with a maximum value less than 50 and then saves the workbook. | Explain the purpose of the IsFiltered property for chart series in Aspose.Cells and how to retrieve the number of filtered series. | Show an alternative method to filter chart series based on their average value instead of the maximum using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartFiltering
{
    // Creates a workbook with three data series, adds a column chart, then evaluates each series' maximum cell value. Series whose maximum does not exceed a defined threshold are hidden by setting the IsFiltered property. The example also shows how to read the filtered series count and save the workbook.
    public class FilterSeriesByThreshold
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data: three series in columns B, C, D
                // -------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");

                // Category labels
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                // Series values
                // Series1 values (some below threshold)
                sheet.Cells["B2"].PutValue(5);
                sheet.Cells["B3"].PutValue(12);
                sheet.Cells["B4"].PutValue(18);
                sheet.Cells["B5"].PutValue(22); // only this exceeds threshold

                // Series2 values (all below threshold)
                sheet.Cells["C2"].PutValue(3);
                sheet.Cells["C3"].PutValue(7);
                sheet.Cells["C4"].PutValue(9);
                sheet.Cells["C5"].PutValue(15);

                // Series3 values (all above threshold)
                sheet.Cells["D2"].PutValue(25);
                sheet.Cells["D3"].PutValue(30);
                sheet.Cells["D4"].PutValue(35);
                sheet.Cells["D5"].PutValue(40);

                // -------------------------------------------------
                // Add a column chart and bind the data
                // -------------------------------------------------
                int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Add each series to the chart
                chart.NSeries.Add("B2:B5", true); // Series1
                chart.NSeries.Add("C2:C5", true); // Series2
                chart.NSeries.Add("D2:D5", true); // Series3

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // -------------------------------------------------
                // Define the threshold and filter series
                // -------------------------------------------------
                double threshold = 20.0;

                // Helper array with the data ranges of each series (same order as added)
                string[] seriesRanges = { "B2:B5", "C2:C5", "D2:D5" };

                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    // Determine the maximum value in the current series range
                    double maxVal = double.MinValue;
                    string[] parts = seriesRanges[i].Split(':');
                    CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);

                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        double val = sheet.Cells[row, area.StartColumn].DoubleValue;
                        if (val > maxVal) maxVal = val;
                    }

                    // If the maximum does not exceed the threshold, hide the series
                    if (maxVal <= threshold)
                    {
                        chart.NSeries[i].IsFiltered = true;
                    }
                }

                // Optional: display how many series are filtered
                Console.WriteLine("Filtered series count: " + chart.FilteredNSeries.Count);

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "FilteredSeriesByThreshold.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                FilterSeriesByThreshold.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }
}
