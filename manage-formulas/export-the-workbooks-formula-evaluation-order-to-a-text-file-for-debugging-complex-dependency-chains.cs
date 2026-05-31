using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Custom monitor that records the order in which cells are calculated
class EvaluationOrderMonitor : AbstractCalculationMonitor
{
    private readonly StreamWriter _writer;

    public EvaluationOrderMonitor(string outputPath)
    {
        // Create (or overwrite) the output text file
        _writer = new StreamWriter(outputPath, false);
    }

    // Called before each cell is calculated
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        // Convert row/column indexes to the usual Excel cell name (e.g., A1)
        string cellName = CellsHelper.CellIndexToName(rowIndex, columnIndex);
        // Write a line indicating the sheet and cell being processed
        _writer.WriteLine($"Sheet{sheetIndex}: {cellName}");
    }

    // Called after each cell is calculated (optional logging)
    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        // No additional action needed for this example
    }

    // Called when a circular reference is detected
    public override bool OnCircular(IEnumerator circularCellsData)
    {
        // Continue processing; returning true tells the engine to ignore the circular reference
        return true;
    }

    // Ensure the writer is properly closed
    public void Close()
    {
        _writer.Flush();
        _writer.Close();
    }
}

class ExportFormulaEvaluationOrder
{
    static void Main(string[] args)
    {
        // Paths can be adjusted as needed
        string inputWorkbookPath = "input.xlsx";
        string outputTextPath = "evaluation_order.txt";

        // Load the workbook (uses the provided load rule)
        Workbook workbook = new Workbook(inputWorkbookPath);

        // Enable the calculation chain so that dependency information is available
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Set up the custom calculation monitor
        var monitor = new EvaluationOrderMonitor(outputTextPath);
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = monitor
        };

        // Perform calculation; the monitor records the evaluation order
        workbook.CalculateFormula(calcOptions);

        // Finalize the output file
        monitor.Close();

        Console.WriteLine($"Formula evaluation order has been exported to '{outputTextPath}'.");
    }
}