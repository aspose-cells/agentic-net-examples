// Title: Cancel Aspose.Cells Workbook.CalculateFormula with a Timeout using CancellationToken
// Description: Demonstrates how to abort a long‑running Workbook.CalculateFormula call in Aspose.Cells by assigning a custom InterruptMonitor that checks a CancellationToken. The token is cancelled after a configurable timeout, causing a CellsException.Interrupted to be thrown, allowing graceful handling and optional saving of a partially calculated workbook.
// Keywords: Aspose.Cells | Workbook.CalculateFormula | CancellationToken | InterruptMonitor | timeout cancellation | cancel formula calculation | CellsException.Interrupted | C# | large workbook | Excel formula abort
// Common Searches: how to cancel Workbook.CalculateFormula in Aspose.Cells | Aspose.Cells interrupt calculation with CancellationToken | timeout for formula calculation Aspose.Cells .NET | catch CellsException.Interrupted after calculation abort | set InterruptMonitor to stop long running Excel calculations
// Developer Intent: Abort a formula calculation in Aspose.Cells when it runs longer than a specified time limit.
// Use Cases: Apply a CancellationTokenInterruptMonitor to stop Workbook.CalculateFormula after a 2‑second timeout and handle the CellsException.Interrupted exception. | Integrate timeout logic into a backend service that processes large Excel files, preventing calculations from blocking resources. | Save a partially calculated workbook after interruption for diagnostics or reporting.
// AI Prompts: Show a C# example that uses Aspose.Cells InterruptMonitor with a CancellationToken to cancel Workbook.CalculateFormula after a configurable timeout. | Explain how to catch CellsException.Interrupted and differentiate it from other Aspose.Cells exceptions when a calculation is aborted. | Provide code to extend CancellationTokenInterruptMonitor to log a message before the calculation is terminated.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// Custom interrupt monitor that checks a CancellationToken
// Demonstrates how to abort a long‑running Workbook.CalculateFormula call in Aspose.Cells by assigning a custom InterruptMonitor that checks a CancellationToken. The token is cancelled after a configurable timeout, causing a CellsException.Interrupted to be thrown, allowing graceful handling and optional saving of a partially calculated workbook.
public class CancellationTokenInterruptMonitor : AbstractInterruptMonitor
{
    private readonly CancellationToken _token;

    public CancellationTokenInterruptMonitor(CancellationToken token)
    {
        _token = token;
    }

    // Return true when the token signals cancellation
    public override bool IsInterruptionRequested => _token.IsCancellationRequested;

    // Let the operation throw a CellsException when interrupted (default behavior)
    public override bool TerminateWithoutException => false;
}

public class CalculateFormulaWithTimeoutDemo
{
    public static void Run()
    {
        // Create a workbook and add some data with formulas
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate a large range to make calculation take noticeable time
        for (int i = 0; i < 5000; i++)
        {
            cells[i, 0].PutValue(i);
            cells[i, 1].Formula = $"=A{i}+10";
        }

        // Set up a cancellation token that will be cancelled after the timeout
        int timeoutMs = 2000; // 2 seconds
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel the token after the specified timeout
            Task.Delay(timeoutMs).ContinueWith(_ => cts.Cancel());

            // Assign the custom interrupt monitor to the workbook
            workbook.InterruptMonitor = new CancellationTokenInterruptMonitor(cts.Token);

            try
            {
                // Perform formula calculation; it will be interrupted if timeout elapses
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed without interruption.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine($"Calculation was interrupted after {timeoutMs} ms.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook (optional, will contain partially calculated results)
            try
            {
                workbook.Save("InterruptedResult.xlsx");
                Console.WriteLine("Workbook saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}

// Entry point for demonstration
class Program
{
    static void Main()
    {
        CalculateFormulaWithTimeoutDemo.Run();
    }
}
