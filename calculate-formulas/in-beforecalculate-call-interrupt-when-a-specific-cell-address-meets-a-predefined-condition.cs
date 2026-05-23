using System;
using System.IO;
using Aspose.Cells;

class MyCalculationMonitor : AbstractCalculationMonitor
{
    private readonly InterruptMonitor _interruptMonitor;
    private readonly int _targetSheet;
    private readonly int _targetRow;
    private readonly int _targetCol;

    public MyCalculationMonitor(InterruptMonitor interruptMonitor, int targetSheet, int targetRow, int targetCol)
    {
        _interruptMonitor = interruptMonitor;
        _targetSheet = targetSheet;
        _targetRow = targetRow;
        _targetCol = targetCol;
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

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and a formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";

            // Set up an interrupt monitor and assign it to the workbook
            InterruptMonitor interruptMonitor = new InterruptMonitor();
            workbook.InterruptMonitor = interruptMonitor;

            // Define the cell that will cause interruption (A3 -> row 2, column 0, zero‑based)
            int targetSheet = 0;
            int targetRow = 2;
            int targetCol = 0;

            // Create a calculation monitor that uses the interrupt monitor
            MyCalculationMonitor calcMonitor = new MyCalculationMonitor(interruptMonitor, targetSheet, targetRow, targetCol);

            // Configure calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = calcMonitor
            };

            try
            {
                // Perform calculation; interruption occurs when A3 is about to be calculated
                workbook.CalculateFormula(options);
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted as expected.");
                // Disable the interrupt monitor before saving to avoid save-time exception
                workbook.InterruptMonitor = null;
            }

            // Save the workbook (partial results may be present)
            string outputPath = "Result.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
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