// Title: C# – Measure Aspose.Cells Formula Calculation Memory With and Without Calculation Chain
// Description: Creates a 5,000‑row, 10‑column workbook of dependent formulas, runs CalculateFormula twice (EnableCalculationChain = false and true), captures GC memory before and after each run, reports the memory delta, and saves both workbooks. Shows the memory impact of the calculation chain on large spreadsheets.
// Keywords: Aspose.Cells | C# | .NET | memory profiling | formula calculation | EnableCalculationChain | large workbook performance | GC.GetTotalMemory | benchmark | Excel formula chain
// Common Searches: Aspose.Cells memory usage formula calculation | EnableCalculationChain performance test | C# measure memory before and after CalculateFormula | benchmark Aspose.Cells formula engine | how to profile memory in large Excel workbook using Aspose
// Developer Intent: The developer wants to compare the memory consumption of Aspose.Cells formula evaluation when the calculation chain is disabled versus enabled on a sizable worksheet.
// Use Cases: Identify the most memory‑efficient EnableCalculationChain setting for processing massive spreadsheets. | Detect memory leaks by tracking GC usage before and after CalculateFormula in long‑running services. | Evaluate file‑size and performance differences between workbooks saved with and without the calculation chain.
// AI Prompts: Write a reusable C# method that logs peak memory usage for Aspose.Cells CalculateFormula with a toggle for EnableCalculationChain. | Provide a step‑by‑step guide to benchmark formula calculation memory across different workbook sizes in Aspose.Cells. | Suggest optimization techniques to lower memory consumption when evaluating formulas on large Excel files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace FormulaChainMemoryMeasurement
{
    // Creates a 5,000‑row, 10‑column workbook of dependent formulas, runs CalculateFormula twice (EnableCalculationChain = false and true), captures GC memory before and after each run, reports the memory delta, and saves both workbooks. Shows the memory impact of the calculation chain on large spreadsheets.
    class Program
    {
        static void Main()
        {
            try
            {
                // Measure memory usage without calculation chain
                Workbook wbNoChain = CreateWorkbookWithFormulas();
                wbNoChain.Settings.FormulaSettings.EnableCalculationChain = false;

                long memoryBeforeNoChain = GC.GetTotalMemory(true);
                wbNoChain.CalculateFormula();
                long memoryAfterNoChain = GC.GetTotalMemory(true);

                Console.WriteLine("=== Calculation Chain Disabled ===");
                Console.WriteLine($"Memory before calculation: {memoryBeforeNoChain:N0} bytes");
                Console.WriteLine($"Memory after calculation : {memoryAfterNoChain:N0} bytes");
                Console.WriteLine($"Memory used by calculation: {memoryAfterNoChain - memoryBeforeNoChain:N0} bytes");

                string noChainPath = "WithoutChain.xlsx";
                wbNoChain.Save(noChainPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(noChainPath)}");

                // Measure memory usage with calculation chain
                Workbook wbWithChain = CreateWorkbookWithFormulas();
                wbWithChain.Settings.FormulaSettings.EnableCalculationChain = true;

                long memoryBeforeWithChain = GC.GetTotalMemory(true);
                wbWithChain.CalculateFormula();
                long memoryAfterWithChain = GC.GetTotalMemory(true);

                Console.WriteLine("\n=== Calculation Chain Enabled ===");
                Console.WriteLine($"Memory before calculation: {memoryBeforeWithChain:N0} bytes");
                Console.WriteLine($"Memory after calculation : {memoryAfterWithChain:N0} bytes");
                Console.WriteLine($"Memory used by calculation: {memoryAfterWithChain - memoryBeforeWithChain:N0} bytes");

                string withChainPath = "WithChain.xlsx";
                wbWithChain.Save(withChainPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(withChainPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Creates a workbook populated with a large number of formulas.
        // Each cell in column B contains a formula that depends on the cell to its left.
        private static Workbook CreateWorkbookWithFormulas()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            int rows = 5000;   // Adjust for desired dataset size
            int cols = 10;

            // Fill column A with numeric values
            for (int r = 0; r < rows; r++)
            {
                cells[r, 0].PutValue(r + 1);
            }

            // Create dependent formulas in subsequent columns
            for (int c = 1; c < cols; c++)
            {
                // Convert previous column index to column letter (e.g., 0 -> "A")
                string prevColLetter = CellsHelper.ColumnIndexToName(c - 1);
                for (int r = 0; r < rows; r++)
                {
                    // Formula: =PreviousColumn{row}*1.01
                    string formula = $"={prevColLetter}{r + 1}*1.01";
                    cells[r, c].Formula = formula;
                }
            }

            return workbook;
        }
    }
}
