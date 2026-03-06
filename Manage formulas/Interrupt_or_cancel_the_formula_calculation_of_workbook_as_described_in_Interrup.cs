using System;
using System.Threading;
using Aspose.Cells;

class CustomInterruptMonitor : AbstractInterruptMonitor
{
    // Flag indicating whether an interruption has been requested
    private volatile bool _requestInterrupt = false;

    public CustomInterruptMonitor()
    {
        // Simulate a delayed interruption request (e.g., after 500 ms)
        new Thread(() =>
        {
            Thread.Sleep(500);
            _requestInterrupt = true;
        }).Start();
    }

    // Called by Aspose.Cells during long‑running operations
    public override bool IsInterruptionRequested => _requestInterrupt;

    // When true, the operation terminates silently; false throws CellsException
    public override bool TerminateWithoutException => false;
}

class Program
{
    static void Main()
    {
        // Path to the XLSX file to be loaded
        string inputFile = "LargeWorkbook.xlsx";

        // Configure load options with the custom interrupt monitor
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = new CustomInterruptMonitor()
        };

        try
        {
            // Load the workbook; the loading process can be interrupted
            Workbook workbook = new Workbook(inputFile, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");

            // Attempt to calculate all formulas; the same monitor applies here
            workbook.CalculateFormula();
            Console.WriteLine("Formula calculation completed.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            // Expected when the interrupt monitor signals an interruption
            Console.WriteLine("Operation was interrupted as requested.");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}