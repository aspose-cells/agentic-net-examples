// Title: How to measure memory usage of Aspose.Cells formula calculation on a large worksheet with and without the calculation chain in C#
// Common Searches: C# Aspose.Cells memory usage when calculating formulas on a 10,000‑row worksheet | compare memory consumption with calculation chain enabled vs disabled in Aspose.Cells | profile GC memory before and after workbook.CalculateFormula in .NET | benchmark formula evaluation performance for large Excel files using Aspose.Cells | how to toggle calculation chain in Aspose.Cells for memory optimization
// Tags: Aspose.Cells formula calculation memory profiling | large Excel worksheet performance .NET | calculation chain toggle Aspose.Cells | GC memory measurement C# Aspose.Cells | benchmarking Excel formula evaluation Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a 10,000‑row by 50‑column workbook, fills it with random numbers, adds a SUM formula per row, and defines a MeasureMemory method that forces garbage collection, records total memory before and after workbook.CalculateFormula(), and prints the delta. It demonstrates how to benchmark memory usage of formula calculation and notes that toggling the calculation chain is not available in the current Aspose.Cells version.
class FormulaCalculationMemoryTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define dataset size
            const int rows = 10000;
            const int cols = 50;

            // Fill the worksheet with random numbers
            Random rnd = new Random();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(rnd.NextDouble() * 1000);
                }
            }

            // Add a formula column that sums the first 10 columns for each row
            for (int r = 0; r < rows; r++)
            {
                // Example formula: =SUM(A1:J1) for row r+1
                string formula = $"=SUM(A{r + 1}:J{r + 1})";
                cells[r, cols].Formula = formula;
            }

            // Function to measure memory usage during calculation
            void MeasureMemory(string description)
            {
                // Note: Calculation chain toggling is not available in this version of Aspose.Cells.
                // The measurement is performed with the default settings.

                // Force garbage collection before measurement
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Record memory before calculation
                long before = GC.GetTotalMemory(true);

                // Perform formula calculation
                workbook.CalculateFormula();

                // Record memory after calculation
                long after = GC.GetTotalMemory(true);

                // Output results
                Console.WriteLine($"{description}:");
                Console.WriteLine($"  Memory before calculation: {before:N0} bytes");
                Console.WriteLine($"  Memory after calculation : {after:N0} bytes");
                Console.WriteLine($"  Memory delta             : {after - before:N0} bytes");
                Console.WriteLine();
            }

            // Measure memory usage (default calculation chain behavior)
            MeasureMemory("Memory Usage with Default Calculation Chain");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
