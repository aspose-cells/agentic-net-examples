using System;
using Aspose.Cells;

class MemoryConsumptionComparison
{
    static void Main()
    {
        // Create two identical large workbooks
        Workbook defaultWb = CreateLargeWorkbook();
        Workbook customWb = CreateLargeWorkbook();

        // Measure memory usage with the default calculation engine
        ForceGarbageCollection();
        long beforeDefault = GC.GetTotalMemory(true);
        defaultWb.CalculateFormula(); // uses built‑in engine
        ForceGarbageCollection();
        long afterDefault = GC.GetTotalMemory(true);
        long usedDefault = afterDefault - beforeDefault;
        Console.WriteLine($"Default engine memory usage: {usedDefault / 1024} KB");

        // Measure memory usage with a custom calculation engine
        CalculationOptions customOptions = new CalculationOptions
        {
            CustomEngine = new PassThroughEngine()
        };

        ForceGarbageCollection();
        long beforeCustom = GC.GetTotalMemory(true);
        customWb.CalculateFormula(customOptions);
        ForceGarbageCollection();
        long afterCustom = GC.GetTotalMemory(true);
        long usedCustom = afterCustom - beforeCustom;
        Console.WriteLine($"Custom engine memory usage: {usedCustom / 1024} KB");

        // Save one of the workbooks (optional, demonstrates lifecycle usage)
        customWb.Save("MemoryComparisonResult.xlsx");
    }

    // Creates a workbook with a large number of formulas to stress the calculation engine
    static Workbook CreateLargeWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        int rows = 2000;   // number of rows
        int cols = 50;     // number of data columns
        int formulaCol = cols; // column where the heavy formula will be placed

        // Fill cells with numeric data
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                cells[i, j].PutValue(i + j);
            }
        }

        // Add a SUM formula per row that references the whole row
        for (int i = 0; i < rows; i++)
        {
            string range = $"A{i + 1}:{GetColumnName(cols)}{i + 1}";
            cells[i, formulaCol].Formula = $"=SUM({range})";
        }

        return wb;
    }

    // Forces a full garbage collection to get a more stable memory measurement
    static void ForceGarbageCollection()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    // Converts a 1‑based column index to an Excel column name (e.g., 1 -> A, 27 -> AA)
    static string GetColumnName(int index)
    {
        int dividend = index;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }

    // Custom calculation engine that does not override any function.
    // By leaving Calculate empty, the default engine handles all calculations.
    class PassThroughEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // No custom processing; default engine will compute the result.
        }
    }
}