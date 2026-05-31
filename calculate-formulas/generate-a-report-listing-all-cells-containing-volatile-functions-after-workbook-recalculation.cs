using System;
using System.Collections.Generic;
using Aspose.Cells;

class VolatileFunctionsReport
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure the calculation chain is enabled for accurate dependency tracking
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Create calculation options and attach a custom monitor
        CalculationOptions options = new CalculationOptions();
        var monitor = new VolatileCalculationMonitor(workbook);
        options.CalculationMonitor = monitor;

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula(options);

        // Retrieve the list of cells that contain volatile functions
        List<string> volatileCells = monitor.VolatileCells;

        // Output the report
        Console.WriteLine("Cells containing volatile functions after recalculation:");
        foreach (string cellName in volatileCells)
        {
            Console.WriteLine(cellName);
        }

        // Save the workbook if any changes need to be persisted
        workbook.Save("output.xlsx");
    }

    // Custom calculation monitor that records cells with volatile functions
    private class VolatileCalculationMonitor : AbstractCalculationMonitor
    {
        private readonly Workbook _workbook;

        // Public list to expose the detected volatile cells
        public List<string> VolatileCells { get; } = new List<string>();

        // Known volatile function names (case‑insensitive)
        private static readonly string[] VolatileFunctions = new[]
        {
            "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
        };

        public VolatileCalculationMonitor(Workbook workbook)
        {
            _workbook = workbook;
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            // Get the cell that has just been calculated
            Worksheet sheet = _workbook.Worksheets[sheetIndex];
            Cell cell = sheet.Cells[rowIndex, colIndex];

            // Only examine cells that actually contain a formula
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                string formulaUpper = cell.Formula.ToUpperInvariant();

                // Check if the formula contains any known volatile function
                foreach (string func in VolatileFunctions)
                {
                    if (formulaUpper.Contains(func))
                    {
                        // Record the cell name (avoid duplicates)
                        if (!VolatileCells.Contains(cell.Name))
                        {
                            VolatileCells.Add(cell.Name);
                        }
                        break;
                    }
                }
            }
        }
    }
}