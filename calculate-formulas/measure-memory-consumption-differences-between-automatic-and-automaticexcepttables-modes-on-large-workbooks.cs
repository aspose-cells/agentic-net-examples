using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryComparison
{
    class Program
    {
        // Size of the test workbook
        const int RowCount = 5000;
        const int ColumnCount = 50;

        static void Main()
        {
            // Prepare a large workbook with sample data
            Workbook wb = CreateLargeWorkbook();

            // Measure memory usage with Automatic calculation mode
            long memoryAutomatic = MeasureMemory(() =>
            {
                wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
                // Force a calculation to materialize any internal structures
                wb.CalculateFormula();
            });

            // Measure memory usage with AutomaticExceptTable calculation mode
            long memoryAutomaticExceptTable = MeasureMemory(() =>
            {
                wb.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;
                wb.CalculateFormula();
            });

            // Output the results
            Console.WriteLine($"Memory after Automatic mode: {memoryAutomatic:N0} bytes");
            Console.WriteLine($"Memory after AutomaticExceptTable mode: {memoryAutomaticExceptTable:N0} bytes");
            Console.WriteLine($"Difference: {Math.Abs(memoryAutomatic - memoryAutomaticExceptTable):N0} bytes");

            // Save the workbook (using the standard save rule)
            wb.Save("LargeWorkbook.xlsx", SaveFormat.Xlsx);

            // Clean up
            wb.Dispose();
        }

        // Creates a workbook filled with numeric data
        static Workbook CreateLargeWorkbook()
        {
            Workbook workbook = new Workbook(); // create workbook (rule)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with simple formulas to engage the calculation engine
            for (int i = 0; i < RowCount; i++)
            {
                for (int j = 0; j < ColumnCount; j++)
                {
                    // Example formula: =ROW()+COLUMN()
                    cells[i, j].Formula = $"=ROW()+COLUMN()";
                }
            }

            return workbook;
        }

        // Executes an action and returns the memory used (in bytes) after the action completes
        static long MeasureMemory(Action action)
        {
            // Ensure a clean baseline
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long before = GC.GetTotalMemory(true);

            // Perform the operation we want to measure
            action();

            // Allow any pending allocations to settle
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long after = GC.GetTotalMemory(true);
            return after - before;
        }
    }
}