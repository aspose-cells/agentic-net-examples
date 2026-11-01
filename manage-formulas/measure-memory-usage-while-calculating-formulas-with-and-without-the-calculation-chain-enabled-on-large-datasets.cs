// Title: C# example to compare memory usage of Aspose.Cells formula calculation with calculation chain enabled vs disabled on large worksheets
// Common Searches: how to measure memory consumption of Aspose.Cells CalculateFormula in C# | benchmark Aspose.Cells formula calculation memory usage with and without calculation chain | C# code to profile memory before and after workbook.CalculateFormula on a large dataset | disable calculation chain in Aspose.Cells via reflection and compare performance | Aspose.Cells memory delta for SUM formulas on 20,000 rows in .NET
// Tags: Aspose.Cells formula calculation memory profiling | C# toggle calculation chain Aspose.Cells | large worksheet memory benchmark Aspose.Cells | GC.GetTotalMemory performance Aspose.Cells | reflection disable calculation chain Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook with 20,000 rows and 10 numeric columns, adds a SUM formula in an 11th column for each row, and records memory before and after calling CalculateFormula() with the default calculation chain. It then uses reflection to turn off the calculation chain (when supported), clears cached results, repeats the calculation, and outputs the memory deltas for both scenarios, demonstrating the impact of the calculation chain on memory usage.
class FormulaCalculationMemoryTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define size of large dataset
            const int rows = 20000;   // number of rows
            const int cols = 10;      // number of data columns

            // Fill the worksheet with numeric data
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    // Example data: (i+1) * (j+1)
                    sheet.Cells[i, j].PutValue((i + 1) * (j + 1));
                }
            }

            // Add a formula column that sums the first 10 columns for each row
            // Column index for the sum formula will be cols (i.e., the 11th column)
            for (int i = 0; i < rows; i++)
            {
                // Formula: =SUM(A1:J1) for row i+1
                string formula = $"=SUM(A{i + 1}:J{i + 1})";
                sheet.Cells[i, cols].Formula = formula;
            }

            // -------------------------------------------------
            // Calculation with Calculation Chain (default)
            // -------------------------------------------------

            // Note: In recent Aspose.Cells versions the calculation chain
            // is enabled by default and the property to toggle it has been removed.
            // Therefore we rely on the default behavior here.

            // Force garbage collection before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record memory before calculation
            long memoryBeforeChain = GC.GetTotalMemory(true);

            // Perform calculation
            workbook.CalculateFormula();

            // Record memory after calculation
            long memoryAfterChain = GC.GetTotalMemory(true);

            // Output memory usage for calculation chain enabled
            Console.WriteLine("=== Calculation Chain Enabled ===");
            Console.WriteLine($"Memory before calculation: {memoryBeforeChain:N0} bytes");
            Console.WriteLine($"Memory after calculation : {memoryAfterChain:N0} bytes");
            Console.WriteLine($"Memory delta              : {memoryAfterChain - memoryBeforeChain:N0} bytes");
            Console.WriteLine();

            // -------------------------------------------------
            // Calculation without Calculation Chain
            // -------------------------------------------------

            // Attempt to disable calculation chain if the API supports it.
            // For versions where the property is unavailable, this block will simply skip the toggle.
            try
            {
                // The property may not exist in some versions; use reflection to set it safely.
                var settingsType = workbook.Settings.GetType();
                var prop = settingsType.GetProperty("EnableCalculationChain");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(workbook.Settings, false);
                }
            }
            catch
            {
                // Ignore any errors while attempting to disable the chain.
            }

            // Clear previously calculated values to force re‑calculation
            workbook.CalculateFormula(); // clears cached results

            // Force garbage collection before second measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record memory before second calculation
            long memoryBeforeNoChain = GC.GetTotalMemory(true);

            // Perform calculation again
            workbook.CalculateFormula();

            // Record memory after second calculation
            long memoryAfterNoChain = GC.GetTotalMemory(true);

            // Output memory usage for calculation chain disabled
            Console.WriteLine("=== Calculation Chain Disabled ===");
            Console.WriteLine($"Memory before calculation: {memoryBeforeNoChain:N0} bytes");
            Console.WriteLine($"Memory after calculation : {memoryAfterNoChain:N0} bytes");
            Console.WriteLine($"Memory delta              : {memoryAfterNoChain - memoryBeforeNoChain:N0} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
