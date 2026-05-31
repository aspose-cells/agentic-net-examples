using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartFilteringDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for two series
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Series 1 values (column B)
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(25);
                sheet.Cells["B4"].PutValue(5);

                // Series 2 values (column C)
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(12);
                sheet.Cells["C4"].PutValue(18);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Set data ranges for the two series
                chart.NSeries.Add("B2:B4", true); // Series1
                chart.NSeries.Add("C2:C4", true); // Series2
                chart.NSeries.CategoryData = "A2:A4";

                // Define the threshold – only series with at least one value greater than this will be shown
                double threshold = 15.0;

                // Iterate through each series and hide those that do NOT exceed the threshold
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    // Retrieve the range string for the series values (e.g., "B2:B4")
                    string range = chart.NSeries[i].Values;

                    // Split the range to obtain the start cell address
                    string startCell = range.Split(':')[0];

                    // Get the first cell of the series to evaluate a representative value
                    Cell firstCell = sheet.Cells[startCell];
                    double cellValue = firstCell.Type == CellValueType.IsNumeric ? firstCell.DoubleValue : 0.0;

                    // If the representative value does not exceed the threshold, filter (hide) the series
                    if (cellValue <= threshold)
                    {
                        chart.NSeries[i].IsFiltered = true;
                    }
                }

                // Optional: display count of filtered series
                Console.WriteLine("Filtered series count: " + chart.FilteredNSeries.Count);

                // Save the workbook
                string outputPath = "ChartFilteredByThreshold.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}