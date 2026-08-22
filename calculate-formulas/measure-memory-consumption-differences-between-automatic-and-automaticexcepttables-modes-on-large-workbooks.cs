// Title: Measure memory consumption of Automatic vs AutomaticExceptTable calculation modes in Aspose.Cells for .NET with a 100k‑row workbook
// AI Prompts: Create a C# console program that builds a 100,000‑row, 10‑column worksheet, assigns simple numeric formulas, forces garbage collection, and records memory usage before and after setting FormulaSettings.CalculationMode to Automatic. | Write code to switch the workbook's FormulaSettings.CalculationMode to AutomaticExceptTable, recalculate formulas, trigger GC, and capture the memory delta for comparison with the Automatic mode. | Generate console output that shows memory in megabytes for workbook population, Automatic mode, AutomaticExceptTable mode, and the differences between them, then save the workbook as an XLSX file.
// Common Searches: aspnet memory profiling Aspose.Cells Automatic calculation mode | compare memory usage of Automatic and AutomaticExceptTable modes in Aspose.Cells | how much memory does Aspose.Cells use when calculating formulas in large Excel files | C# benchmark Aspose.Cells formula calculation memory consumption | measure .NET GC impact after workbook.CalculateFormula with Aspose.Cells
// Tags: automatic calculation mode memory profiling Aspose.Cells | automaticexcepttable mode memory benchmark Aspose.Cells | large workbook formula evaluation memory .NET | Aspose.Cells formula settings performance comparison | C# garbage collection memory measurement workbook

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryComparison
{
    // // Builds a 100,000‑row worksheet with simple formulas, records memory after population, switches FormulaSettings.CalculationMode to Automatic and AutomaticExceptTable, recalculates, forces GC, logs memory increases for each mode, and saves the workbook as LargeWorkbookMemoryComparison.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define size of a large dataset
            const int totalRows = 100_000;   // 100k rows
            const int totalCols = 10;        // 10 columns

            // Fill the worksheet with sample data
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Simple numeric formula to involve calculation engine
                    cells[row, col].Formula = $"={row}+{col}";
                }
            }

            // Force a full garbage collection before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Measure memory after workbook creation and data population
            long memoryAfterPopulation = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory after populating workbook: {memoryAfterPopulation / (1024 * 1024)} MB");

            // -----------------------------------------------------------------
            // Scenario 1: Automatic calculation mode
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Optionally trigger calculation (Aspose.Cells does not respect the mode,
            // but we include it to keep the scenario realistic)
            workbook.CalculateFormula();

            // Measure memory after setting Automatic mode and calculation
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long memoryAfterAutomatic = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory after Automatic mode: {memoryAfterAutomatic / (1024 * 1024)} MB");

            // -----------------------------------------------------------------
            // Scenario 2: AutomaticExceptTable calculation mode
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Re‑calculate formulas to reflect the new mode
            workbook.CalculateFormula();

            // Measure memory after setting AutomaticExceptTable mode and calculation
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long memoryAfterAutomaticExceptTable = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory after AutomaticExceptTable mode: {memoryAfterAutomaticExceptTable / (1024 * 1024)} MB");

            // -----------------------------------------------------------------
            // Output the differences
            // -----------------------------------------------------------------
            long diffAutomatic = memoryAfterAutomatic - memoryAfterPopulation;
            long diffExceptTable = memoryAfterAutomaticExceptTable - memoryAfterPopulation;

            Console.WriteLine($"Memory increase (Automatic): {diffAutomatic / (1024 * 1024)} MB");
            Console.WriteLine($"Memory increase (AutomaticExceptTable): {diffExceptTable / (1024 * 1024)} MB");

            // Save the workbook to verify that saving works (save rule)
            workbook.Save("LargeWorkbookMemoryComparison.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as LargeWorkbookMemoryComparison.xlsx");
        }
    }
}
