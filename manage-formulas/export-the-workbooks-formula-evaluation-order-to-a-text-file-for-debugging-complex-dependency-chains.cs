// Title: Export Formula Evaluation Order with Aspose.Cells (C#) for Debugging
// Description: Shows how to enable the calculation chain, attach a custom AbstractCalculationMonitor, run CalculateFormula, and write the resulting cell‑calculation sequence to a text file using Aspose.Cells.
// Keywords: Aspose.Cells | C# | formula evaluation order | calculation monitor | AbstractCalculationMonitor | EnableCalculationChain | export to text | debug formulas | dependency chain | workbook calculation logging
// Common Searches: Aspose.Cells get formula calculation order | log cell evaluation sequence Aspose | export calculation monitor output to file | enable calculation chain for debugging | write formula evaluation order C#
// Developer Intent: Create a text file that lists the exact order in which formulas are evaluated in an Aspose.Cells workbook.
// Use Cases: Diagnose complex inter‑dependent formulas and locate circular references. | Generate audit logs of calculation steps for performance analysis. | Validate expected calculation order in automated unit tests. | Document the calculation flow for end‑users or support teams.
// AI Prompts: Add timestamps to each entry in the evaluation log. | Convert the exported FormulaEvaluationOrder.txt into a GraphViz dependency diagram. | Show how to disable EnableCalculationChain after debugging to improve runtime speed. | Explain how to capture the evaluation order for a single worksheet only. | Provide a PowerShell script that parses FormulaEvaluationOrder.txt and summarizes cell counts per sheet.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Custom monitor to capture the order in which cells are calculated
// Shows how to enable the calculation chain, attach a custom AbstractCalculationMonitor, run CalculateFormula, and write the resulting cell‑calculation sequence to a text file using Aspose.Cells.
class FormulaEvaluationMonitor : AbstractCalculationMonitor
{
    // List to store cell addresses in the order they are processed
    public List<string> EvaluationOrder { get; } = new List<string>();

    // Called before each cell calculation
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        // Convert row/column indices to Excel cell name (e.g., A1)
        string cellName = CellsHelper.CellIndexToName(rowIndex, columnIndex);
        EvaluationOrder.Add($"{sheetIndex}:{cellName}");
    }

    // AfterCalculate can be left empty or used for additional logging
    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex) { }
}

class ExportFormulaEvaluationOrder
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a workbook and add sample data/formulas
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample values
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].Formula = "=A1+A2";          // Depends on A1 and A2
        cells["B1"].Formula = "=A3*2";           // Depends on A3
        cells["B2"].Formula = "=SUM(A1:A3)";     // Depends on A1, A2, A3

        // -------------------------------------------------
        // 2. Enable calculation chain (required for monitoring)
        // -------------------------------------------------
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // -------------------------------------------------
        // 3. Set up calculation options with the custom monitor
        // -------------------------------------------------
        var monitor = new FormulaEvaluationMonitor();
        var calcOptions = new CalculationOptions
        {
            CalculationMonitor = monitor
        };

        // -------------------------------------------------
        // 4. Perform calculation (this will populate the monitor)
        // -------------------------------------------------
        workbook.CalculateFormula(calcOptions);

        // -------------------------------------------------
        // 5. Export the captured evaluation order to a text file
        // -------------------------------------------------
        string outputPath = "FormulaEvaluationOrder.txt";
        File.WriteAllLines(outputPath, monitor.EvaluationOrder);

        Console.WriteLine($"Formula evaluation order exported to: {outputPath}");
    }
}
