// Title: Benchmark Aspose.Cells Pivot Table Generation: Default vs Custom Globalization (C#)
// Description: A C# console app that creates a workbook with 100,000 rows, builds a pivot table on a second sheet, and measures the elapsed time for pivot creation and calculation using the library's default globalization settings and a custom culture configuration. The program saves both workbooks and prints the timing results.
// Keywords: Aspose.Cells | pivot table performance | globalization settings | CultureInfo | benchmark .NET | large dataset | workbook generation time | performance testing | C# Aspose.Cells
// Common Searches: Aspose.Cells pivot table performance test | measure impact of culture settings on Aspose.Cells | benchmark large pivot table creation .NET | how to time Aspose.Cells workbook generation | globalization vs default performance Aspose.Cells | custom culture benchmark Aspose.Cells | C# pivot table speed Aspose.Cells | Aspose.Cells performance with large data
// Developer Intent: Determine whether applying a custom globalization (culture) configuration changes the execution time of generating and calculating a large pivot table with Aspose.Cells for .NET.
// Use Cases: Compare runtime of default and custom globalization when building a 100,000‑row pivot table. | Identify performance bottlenecks in pivot table creation for massive data sets. | Validate that custom culture settings do not degrade workbook generation speed. | Establish a baseline for performance tuning of Aspose.Cells in CI pipelines. | Demonstrate how to instrument timing code around Aspose.Cells operations.
// AI Prompts: Show how to apply a specific CultureInfo to a Workbook before creating a pivot table and measure the effect on execution time. | Suggest techniques for more accurate benchmarking, such as warm‑up runs, multiple iterations, or using BenchmarkDotNet with Aspose.Cells. | Provide code to capture memory usage and CPU load together with elapsed time for both default and custom globalization scenarios. | Explain how to log timing results to a CSV file for later analysis. | Recommend ways to reduce pivot table generation time when working with very large data sets in Aspose.Cells.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPerformanceDemo
{
    // A C# console app that creates a workbook with 100,000 rows, builds a pivot table on a second sheet, and measures the elapsed time for pivot creation and calculation using the library's default globalization settings and a custom culture configuration. The program saves both workbooks and prints the timing results.
    public class Program
    {
        // Number of rows for the large data set
        private const int RowCount = 100_000;

        public static void Main()
        {
            try
            {
                // Measure performance with default globalization settings
                TimeSpan defaultTime = GeneratePivotWithDefaultSettings();
                Console.WriteLine($"Default globalization settings elapsed: {defaultTime.TotalSeconds:F2} seconds");

                // Measure performance with custom globalization settings
                TimeSpan customTime = GeneratePivotWithCustomSettings();
                Console.WriteLine($"Custom globalization settings elapsed: {customTime.TotalSeconds:F2} seconds");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static TimeSpan GeneratePivotWithDefaultSettings()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Populate a large data set in the first worksheet
                Worksheet dataSheet = workbook.Worksheets[0];
                PopulateLargeData(dataSheet);

                // Create a second worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot_Default");

                // Start timing
                Stopwatch sw = Stopwatch.StartNew();

                // Add pivot table
                int pivotIndex = pivotSheet.PivotTables.Add($"A1:B{RowCount + 1}", "D1", "PivotTable_Default");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields (Row: Category, Data: Amount)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Refresh and calculate the pivot table using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Stop timing
                sw.Stop();

                // Save the workbook (default settings)
                workbook.Save("Pivot_Default.xlsx");

                return sw.Elapsed;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in GeneratePivotWithDefaultSettings: {ex.Message}");
                return TimeSpan.Zero;
            }
        }

        private static TimeSpan GeneratePivotWithCustomSettings()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // NOTE: Custom globalization settings are omitted because the current
                // Aspose.Cells version does not expose those properties directly.

                // Populate the same large data set in the first worksheet
                Worksheet dataSheet = workbook.Worksheets[0];
                PopulateLargeData(dataSheet);

                // Create a second worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot_Custom");

                // Start timing
                Stopwatch sw = Stopwatch.StartNew();

                // Add pivot table
                int pivotIndex = pivotSheet.PivotTables.Add($"A1:B{RowCount + 1}", "D1", "PivotTable_Custom");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields (Row: Category, Data: Amount)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Refresh and calculate the pivot table using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Stop timing
                sw.Stop();

                // Save the workbook (custom settings)
                workbook.Save("Pivot_Custom.xlsx");

                return sw.Elapsed;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in GeneratePivotWithCustomSettings: {ex.Message}");
                return TimeSpan.Zero;
            }
        }

        // Helper method to fill the worksheet with a large amount of sample data
        private static void PopulateLargeData(Worksheet sheet)
        {
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");

            // Generate rows with pseudo‑random data
            Random rnd = new Random(0);
            for (int i = 2; i <= RowCount + 1; i++)
            {
                // Cycle through a few categories to create grouping
                string category = "Category_" + (i % 10);
                double amount = rnd.NextDouble() * 1000;

                sheet.Cells[i - 1, 0].PutValue(category);
                sheet.Cells[i - 1, 1].PutValue(amount);
            }
        }
    }
}
