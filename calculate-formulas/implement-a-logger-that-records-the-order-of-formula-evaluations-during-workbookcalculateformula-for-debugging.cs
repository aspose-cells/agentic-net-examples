// Title: Log Formula Evaluation Order with a Custom AbstractCalculationMonitor in Aspose.Cells for .NET
// Description: Demonstrates how to create a FormulaEvaluationLogger that inherits from AbstractCalculationMonitor, captures each cell address after calculation, and attaches it to CalculationOptions for Workbook.CalculateFormula. Use the logger to debug formula dependencies, track evaluation sequence, and identify performance issues in C# projects using Aspose.Cells.
// Keywords: Aspose.Cells | AbstractCalculationMonitor | Formula evaluation logger | Workbook.CalculateFormula | C# | .NET | track cell calculation order | debug formula dependencies | custom calculation monitor | volatile function handling
// Common Searches: Aspose.Cells log formula evaluation order | custom AbstractCalculationMonitor example C# | track cell calculation sequence Aspose.Cells | how to monitor Workbook.CalculateFormula | debug formula dependencies Aspose.Cells
// Developer Intent: Create a reusable logger that records the exact order cells are evaluated during Workbook.CalculateFormula for debugging and performance analysis.
// Use Cases: Debug complex workbooks by reviewing the precise formula evaluation sequence. | Detect performance bottlenecks caused by volatile functions or deep dependency chains. | Validate recursive calculation behavior in large spreadsheets. | Export the evaluation order for audit or reporting purposes.
// AI Prompts: Generate a C# AbstractCalculationMonitor that logs cell addresses with timestamps to a text file. | Show how to modify the logger to output the evaluation order as a CSV with sheet name, cell address, and calculation time. | Provide code to filter logged entries so only cells containing volatile functions (e.g., NOW, RAND) are recorded.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to create a FormulaEvaluationLogger that inherits from AbstractCalculationMonitor, captures each cell address after calculation, and attaches it to CalculationOptions for Workbook.CalculateFormula. Use the logger to debug formula dependencies, track evaluation sequence, and identify performance issues in C# projects using Aspose.Cells.
class FormulaEvaluationLogger : AbstractCalculationMonitor
{
    // Stores the order in which cells are calculated
    private readonly List<string> _evaluationOrder = new List<string>();
    public IReadOnlyList<string> EvaluationOrder => _evaluationOrder.AsReadOnly();

    // Called before a cell is calculated (optional logging can be added here)
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        // No action needed for this logger
    }

    // Called after a cell has been calculated – record its address
    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        string cellAddress = CellsHelper.CellIndexToName(rowIndex, columnIndex);
        _evaluationOrder.Add($"Sheet{sheetIndex}!{cellAddress}");
    }

    // Handle circular references – simply continue calculation
    public override bool OnCircular(System.Collections.IEnumerator circularCellsData)
    {
        return true; // Continue calculation
    }

    // Helper to output the logged order
    public void PrintLog()
    {
        foreach (var entry in _evaluationOrder)
        {
            Console.WriteLine(entry);
        }
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and add some formulas
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].Formula = "=1+2";
        sheet.Cells["A2"].Formula = "=A1*3";
        sheet.Cells["A3"].Formula = "=SUM(A1:A2)";
        sheet.Cells["B1"].Formula = "=NOW()"; // volatile function

        // Instantiate the custom calculation monitor
        var logger = new FormulaEvaluationLogger();

        // Configure calculation options to use the monitor
        CalculationOptions options = new CalculationOptions
        {
            CalculationMonitor = logger,
            Recursive = true,
            IgnoreError = false
        };

        // Perform formula calculation with monitoring
        workbook.CalculateFormula(options);

        // Display calculated values
        Console.WriteLine("Calculated Values:");
        Console.WriteLine($"A1 = {sheet.Cells["A1"].Value}");
        Console.WriteLine($"A2 = {sheet.Cells["A2"].Value}");
        Console.WriteLine($"A3 = {sheet.Cells["A3"].Value}");
        Console.WriteLine($"B1 = {sheet.Cells["B1"].Value}");

        // Output the order in which formulas were evaluated
        Console.WriteLine("\nFormula Evaluation Order:");
        logger.PrintLog();

        // Save the workbook (optional)
        workbook.Save("LoggedWorkbook.xlsx");
    }
}
