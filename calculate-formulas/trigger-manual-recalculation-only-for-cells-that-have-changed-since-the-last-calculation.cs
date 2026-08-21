// Title: Aspose.Cells for .NET – Manual Calculation with ChangeMonitor to Recalculate Only Modified Cells
// Description: Demonstrates how to set Aspose.Cells to manual calculation mode, use a custom AbstractCalculationMonitor (ChangeMonitor) to log cells whose values actually change, update a source cell, trigger selective recalculation of dependent formulas, and save the workbook.
// Keywords: Aspose.Cells manual calculation | C# Aspose.Cells ChangeMonitor | recalculate only changed cells .NET | AbstractCalculationMonitor example | track cell value changes Aspose.Cells | selective formula recalculation | Aspose.Cells CalculationOptions | Workbook.CalculateFormula manual mode | Aspose.Cells sample code GitHub
// Common Searches: Aspose.Cells manual calculation mode C# example | How to recalculate only changed cells with Aspose.Cells | Aspose.Cells ChangeMonitor to detect updated cells | Selective formula recalculation Aspose.Cells .NET | Custom CalculationMonitor Aspose.Cells tutorial
// Developer Intent: The developer wants to perform manual formula recalculation, detect which cells changed after each calculation, and update only those dependent cells.
// Use Cases: Switch workbook to manual calculation (CalcModeType.Manual) and invoke Workbook.CalculateFormula only when needed. | Implement a subclass of AbstractCalculationMonitor to capture the A1 reference, original value, and new value of each cell that changes during calculation. | Update one or more source cells, run CalculateFormula with a CalculationMonitor, and let Aspose.Cells automatically refresh only the affected dependent cells. | Log changed cells for auditing or debugging, then save the workbook with the updated results.
// AI Prompts: Generate C# code that configures Aspose.Cells for manual calculation, updates a source cell, and uses a custom ChangeMonitor to log only cells whose values changed. | Explain how Aspose.Cells determines which cells need recalculation in manual mode and how to retrieve changed cell references via CalculationOptions and a CalculationMonitor. | Provide a step‑by‑step guide to modify multiple source cells, trigger selective recalculation, log changed cells, and save the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using System.Collections;

// Demonstrates how to set Aspose.Cells to manual calculation mode, use a custom AbstractCalculationMonitor (ChangeMonitor) to log cells whose values actually change, update a source cell, trigger selective recalculation of dependent formulas, and save the workbook.
public class ManualRecalculationDemo
{
    // Custom monitor to report cells whose value actually changed after a calculation
    private class ChangeMonitor : AbstractCalculationMonitor
    {
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            if (ValueChanged)
            {
                Console.WriteLine($"Cell {CellReference(rowIndex, colIndex)} changed from [{OriginalValue}] to [{CalculatedValue}]");
            }
        }

        private string CellReference(int row, int col)
        {
            // Convert zero‑based row/col to A1 style reference
            return CellsHelper.CellIndexToName(row, col);
        }
    }

    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new workbook and set manual calculation mode
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Manual mode ensures that calculations are performed only when we explicitly call them
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // -------------------------------------------------
        // 2. Populate initial data and formulas
        // -------------------------------------------------
        cells["A1"].PutValue(10);               // source value
        cells["A2"].PutValue(20);               // source value
        cells["B1"].Formula = "=A1+A2";         // depends on A1 and A2
        cells["C1"].Formula = "=B1*2";          // depends on B1

        // -------------------------------------------------
        // 3. First full calculation
        // -------------------------------------------------
        CalculationOptions firstCalcOpts = new CalculationOptions
        {
            CalculationMonitor = new ChangeMonitor()
        };
        workbook.CalculateFormula(firstCalcOpts);

        Console.WriteLine($"After first calculation: B1 = {cells["B1"].Value}, C1 = {cells["C1"].Value}");

        // -------------------------------------------------
        // 4. Modify only one source cell (A1)
        // -------------------------------------------------
        cells["A1"].PutValue(30);   // change triggers dependent recalculation

        // -------------------------------------------------
        // 5. Recalculate – only cells that depend on changed cells are updated
        // -------------------------------------------------
        CalculationOptions secondCalcOpts = new CalculationOptions
        {
            CalculationMonitor = new ChangeMonitor()
        };
        workbook.CalculateFormula(secondCalcOpts);

        Console.WriteLine($"After changing A1: B1 = {cells["B1"].Value}, C1 = {cells["C1"].Value}");

        // -------------------------------------------------
        // 6. Save the workbook
        // -------------------------------------------------
        workbook.Save("ManualRecalculationDemo.xlsx");
    }
}
