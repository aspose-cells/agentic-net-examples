// Title: Benchmark Aspose.Cells pivot table creation time with default versus French (fr-FR) globalization in C#
// Common Searches: how to benchmark Aspose.Cells pivot table generation with custom CultureInfo in C# | measure performance impact of setting workbook.Settings.CultureInfo in Aspose.Cells | compare default and fr-FR globalization speed for large pivot tables using Aspose.Cells | C# code to time pivot table creation and calculation with Aspose.Cells Stopwatch | performance test for Aspose.Cells pivot tables on a 100,000‑row dataset
// Tags: benchmarking Aspose.Cells pivot table creation | custom CultureInfo effect on Aspose.Cells performance | large dataset pivot table timing .NET | Aspose.Cells globalization settings impact | pivot table calculation speed C#

using System;
using System.Diagnostics;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotFieldType

// The example creates a workbook with 100,000 rows of mixed data, then measures and prints the elapsed time for creating and calculating a pivot table using default globalization settings versus after applying a French (fr-FR) CultureInfo, providing a direct performance comparison of the two globalization configurations.
class PivotTablePerformanceTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Generate a large data set (e.g., 100,000 rows, 5 columns)
            const int rowCount = 100000;
            const int colCount = 5;

            // Add header row
            for (int c = 0; c < colCount; c++)
            {
                sheet.Cells[0, c].PutValue($"Column{c + 1}");
            }

            // Fill data rows with random numbers and dates
            Random rnd = new Random();
            for (int r = 1; r <= rowCount; r++)
            {
                sheet.Cells[r, 0].PutValue(rnd.Next(1, 1000)); // Integer
                sheet.Cells[r, 1].PutValue(rnd.NextDouble() * 1000); // Double
                sheet.Cells[r, 2].PutValue(DateTime.Today.AddDays(rnd.Next(-365, 365))); // Date
                sheet.Cells[r, 3].PutValue($"Category{rnd.Next(1, 10)}"); // Text
                sheet.Cells[r, 4].PutValue(rnd.Next(0, 2) == 0); // Boolean
            }

            // Define the data range for the pivot table
            string dataRange = $"A1:E{rowCount + 1}";

            // Measure performance with default globalization settings
            Stopwatch swDefault = Stopwatch.StartNew();

            // Add pivot table using default settings (Add returns the index)
            int pivotIndexDefault = sheet.PivotTables.Add($"=Sheet1!{dataRange}", "H1", "PivotDefault");
            PivotTable pivotDefault = sheet.PivotTables[pivotIndexDefault];

            // Configure pivot fields (example: sum of Column1, average of Column2)
            pivotDefault.AddFieldToArea(PivotFieldType.Row, 3);   // Category column as row field
            pivotDefault.AddFieldToArea(PivotFieldType.Data, 0);  // Column1 as data field (sum)
            pivotDefault.AddFieldToArea(PivotFieldType.Data, 1);  // Column2 as data field (average)

            // Refresh and calculate the pivot table
            pivotDefault.RefreshData();
            pivotDefault.CalculateData();

            swDefault.Stop();

            // Apply custom culture (e.g., French - France) to the workbook
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Measure performance with custom globalization settings
            Stopwatch swCustom = Stopwatch.StartNew();

            // Add another pivot table after culture change
            int pivotIndexCustom = sheet.PivotTables.Add($"=Sheet1!{dataRange}", "H20", "PivotCustom");
            PivotTable pivotCustom = sheet.PivotTables[pivotIndexCustom];
            pivotCustom.AddFieldToArea(PivotFieldType.Row, 3);
            pivotCustom.AddFieldToArea(PivotFieldType.Data, 0);
            pivotCustom.AddFieldToArea(PivotFieldType.Data, 1);

            // Refresh and calculate the second pivot table
            pivotCustom.RefreshData();
            pivotCustom.CalculateData();

            swCustom.Stop();

            // Output the measured times
            Console.WriteLine($"Pivot table creation time with default globalization: {swDefault.ElapsedMilliseconds} ms");
            Console.WriteLine($"Pivot table creation time with custom globalization (fr-FR): {swCustom.ElapsedMilliseconds} ms");

            // Optional: Save the workbook to verify results (not required for performance measurement)
            // workbook.Save("PivotPerformanceTest.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
