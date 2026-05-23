using System;
using System.Diagnostics;
using Aspose.Cells;

class MemoryComparison
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define a large dataset size
        int totalRows = 5000;
        int totalCols = 50;
        Random random = new Random();

        // Populate the worksheet with random numeric data
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalCols; col++)
            {
                cells[row, col].PutValue(random.NextDouble());
            }
        }

        // Ensure a clean state before measuring memory
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // -------------------- Automatic mode --------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
        // Optionally trigger a calculation to reflect the mode
        workbook.CalculateFormula();

        long memoryAutomatic = GetCurrentProcessMemory();

        // -------------------- AutomaticExceptTable mode --------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;
        // Recalculate after changing the mode
        workbook.CalculateFormula();

        long memoryAutomaticExceptTable = GetCurrentProcessMemory();

        // Output the memory consumption results
        Console.WriteLine($"Memory (Automatic): {memoryAutomatic / 1024.0 / 1024.0:F2} MB");
        Console.WriteLine($"Memory (AutomaticExceptTable): {memoryAutomaticExceptTable / 1024.0 / 1024.0:F2} MB");

        // Save the workbook (save rule)
        workbook.Save("LargeWorkbook.xlsx");
    }

    // Helper method to retrieve the current process private memory size
    static long GetCurrentProcessMemory()
    {
        Process proc = Process.GetCurrentProcess();
        proc.Refresh();
        return proc.PrivateMemorySize64;
    }
}