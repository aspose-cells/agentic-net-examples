// Title: C# – Interrupt Formula Calculation for a Specific Cell Using Aspose.Cells CalculationMonitor
// Description: This example shows how to create a custom ConditionalInterruptMonitor that inherits AbstractCalculationMonitor and overrides BeforeCalculate. When the monitor detects the target sheet, row, and column, it calls Interrupt() on an InterruptMonitor, causing the workbook calculation to stop. The demo builds a workbook, adds data and a formula in B2, assigns the interrupt monitor via CalculationOptions, catches the Interrupted CellsException, clears the monitor, and optionally saves the partially calculated file.
// Keywords: Aspose.Cells InterruptMonitor | CalculationMonitor C# | BeforeCalculate interrupt | stop formula calculation specific cell | Aspose.Cells custom monitor example | handle Interrupted exception | conditional calculation break
// Common Searches: how to stop Aspose.Cells calculation for one cell | using CalculationMonitor to interrupt formula evaluation in .NET | example of BeforeCalculate interrupting B2 in Aspose.Cells | Aspose.Cells cancel calculation when condition met | C# interrupt workbook calculation with InterruptMonitor
// Developer Intent: Stop the workbook’s formula evaluation when a designated cell is about to be calculated.
// Use Cases: Skip expensive formulas for cells that meet a predefined condition during large‑scale calculations. | Implement a cancellation or timeout feature by halting calculation when a specific cell is reached. | Prevent evaluation of cells that depend on unavailable external data by interrupting in BeforeCalculate.
// AI Prompts: Generate C# code that uses Aspose.Cells InterruptMonitor to cancel calculation when cell C5 exceeds a given threshold. | Explain how to reset the interrupt flag after catching the Interrupted exception and then continue remaining calculations. | Create a version of ConditionalInterruptMonitor that logs the cell address before calling Interrupt() in BeforeCalculate.

using Aspose.Cells;
using System;

// This example shows how to create a custom ConditionalInterruptMonitor that inherits AbstractCalculationMonitor and overrides BeforeCalculate. When the monitor detects the target sheet, row, and column, it calls Interrupt() on an InterruptMonitor, causing the workbook calculation to stop. The demo builds a workbook, adds data and a formula in B2, assigns the interrupt monitor via CalculationOptions, catches the Interrupted CellsException, clears the monitor, and optionally saves the partially calculated file.
public class ConditionalInterruptMonitor : AbstractCalculationMonitor
{
    private readonly InterruptMonitor _interruptMonitor;
    private readonly int _targetSheet;
    private readonly int _targetRow;
    private readonly int _targetCol;

    public ConditionalInterruptMonitor(InterruptMonitor interruptMonitor, int sheetIndex, int rowIndex, int colIndex)
    {
        _interruptMonitor = interruptMonitor;
        _targetSheet = sheetIndex;
        _targetRow = rowIndex;
        _targetCol = colIndex;
    }

    public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
    {
        // Interrupt when the specified cell is about to be calculated
        if (sheetIndex == _targetSheet && rowIndex == _targetRow && colIndex == _targetCol)
        {
            _interruptMonitor.Interrupt();
        }
    }
}

public class Demo
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and a formula in B2 (row 1, column 1)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B2"].Formula = "=A1+A2";

            // Create an interrupt monitor and assign it to the workbook
            InterruptMonitor interruptMonitor = new InterruptMonitor();
            workbook.InterruptMonitor = interruptMonitor;

            // Create a calculation monitor that interrupts when B2 is about to be calculated
            ConditionalInterruptMonitor calcMonitor = new ConditionalInterruptMonitor(
                interruptMonitor,
                sheet.Index,
                1, // row index for B2 (zero‑based)
                1  // column index for B2 (zero‑based)
            );

            // Set calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = calcMonitor
            };

            try
            {
                // Perform calculation; it will be interrupted in BeforeCalculate
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed without interruption.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted as expected.");
                // Clear the interrupt flag before further operations
                workbook.InterruptMonitor = null;
            }

            // Save the workbook (optional, may contain partial results)
            try
            {
                workbook.Save("ConditionalInterruptDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
