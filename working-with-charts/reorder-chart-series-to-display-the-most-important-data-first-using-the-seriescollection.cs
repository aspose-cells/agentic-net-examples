using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesReorder
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for three series
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // Series 1 values
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2 values
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(40);
                sheet.Cells["C3"].PutValue(15);
                sheet.Cells["C4"].PutValue(25);

                // Series 3 values
                sheet.Cells["D1"].PutValue("Series 3");
                sheet.Cells["D2"].PutValue(5);
                sheet.Cells["D3"].PutValue(35);
                sheet.Cells["D4"].PutValue(45);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add all three series to the chart (by column)
                chart.NSeries.Add("B1:D4", true);
                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // Get the series collection
                SeriesCollection seriesColl = chart.NSeries;

                // Reorder series so that the series with the highest total value appears first
                // Simple selection sort using SwapSeries
                int count = seriesColl.Count;
                for (int i = 0; i < count - 1; i++)
                {
                    // Find index of series with maximum sum from i to end
                    int maxIdx = i;
                    double maxSum = GetSeriesSum(seriesColl[i]);

                    for (int j = i + 1; j < count; j++)
                    {
                        double sum = GetSeriesSum(seriesColl[j]);
                        if (sum > maxSum)
                        {
                            maxSum = sum;
                            maxIdx = j;
                        }
                    }

                    // If the maximum is not already at position i, swap them
                    if (maxIdx != i)
                    {
                        seriesColl.SwapSeries(i, maxIdx);
                    }
                }

                // Save the workbook
                string outputPath = "ReorderedSeriesChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to calculate the sum of all point values in a series
        private static double GetSeriesSum(Series series)
        {
            double sum = 0;
            // PointValues returns an array of ChartDataValue objects
            ChartDataValue[] values = series.PointValues;
            if (values != null)
            {
                foreach (ChartDataValue v in values)
                {
                    // Use DoubleValue to get the numeric representation
                    sum += v.DoubleValue;
                }
            }
            return sum;
        }
    }
}