// Title: Cancel Aspose.Cells CalculateFormula with a CancellationToken and custom InterruptMonitor (C#)
// Description: Demonstrates how to abort a long‑running Workbook.CalculateFormula by attaching a TokenInterruptMonitor (derived from AbstractInterruptMonitor) to a workbook, using a CancellationToken that is triggered from a background task.
// Keywords: Aspose.Cells CancellationToken | InterruptMonitor C# | Abort CalculateFormula | Cancel long running formula calculation | CellsException Interrupted | Aspose.Cells multithreading | Workbook.CalculateFormula cancel | C# Excel library cancellation
// Common Searches: how to stop Aspose.Cells CalculateFormula with CancellationToken | Aspose.Cells interrupt monitor example | cancel long running workbook calculation C# | use CancellationToken to abort Excel formula evaluation | Aspose.Cells custom InterruptMonitor usage
// Developer Intent: Implement a mechanism that lets a user or service abort a lengthy CalculateFormula operation in Aspose.Cells.
// Use Cases: Add a UI Cancel button that stops formula calculation on demand. | Enforce a maximum execution time for server‑side workbook processing to prevent timeouts. | Gracefully shut down a background service by cancelling ongoing calculations.
// AI Prompts: Generate a WinForms sample that wires a Cancel button to a CancellationTokenSource and uses TokenInterruptMonitor to stop Workbook.CalculateFormula. | Show how to catch CellsException with code Interrupted and log the cancellation event in an ASP.NET Core API. | Create a reusable helper method that configures a Workbook with a CancellationToken, runs CalculateFormula, and handles interruption exceptions.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsCancellationDemo
{
    // Custom interrupt monitor that checks a CancellationToken
    // Demonstrates how to abort a long‑running Workbook.CalculateFormula by attaching a TokenInterruptMonitor (derived from AbstractInterruptMonitor) to a workbook, using a CancellationToken that is triggered from a background task.
    public class TokenInterruptMonitor : AbstractInterruptMonitor
    {
        private readonly CancellationToken _token;

        public TokenInterruptMonitor(CancellationToken token)
        {
            _token = token;
        }

        // Return true when the token is cancelled – this will cause Aspose.Cells to interrupt the operation
        public override bool IsInterruptionRequested => _token.IsCancellationRequested;

        // Keep default behavior (throw CellsException when interrupted)
        public override bool TerminateWithoutException => false;
    }

    class Program
    {
        static void Main()
        {
            // Prepare a cancellation token source
            var cts = new CancellationTokenSource();

            // Create a workbook and fill it with sample data and formulas
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate many rows to make calculation take noticeable time
            for (int i = 0; i < 50000; i++)
            {
                cells[i, 0].PutValue(i);               // Column A
                cells[i, 1].Formula = $"=A{i + 1}*2";   // Column B depends on A
            }

            // Set a formula that sums a large range (adds more load)
            cells[0, 2].Formula = $"=SUM(B1:B50000)";

            // Attach the custom interrupt monitor to the workbook
            workbook.InterruptMonitor = new TokenInterruptMonitor(cts.Token);

            // Start a background task that will request cancellation after a short delay
            Task.Run(() =>
            {
                Thread.Sleep(2000); // wait 2 seconds
                Console.WriteLine("Cancellation requested.");
                cts.Cancel();       // signal cancellation
            });

            try
            {
                Console.WriteLine("Starting calculation...");
                // Perform calculation; it will be interrupted when the token is set
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted as requested.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook (if calculation finished, otherwise partial results may be saved)
            try
            {
                workbook.Save("CancellationDemo.xlsx");
                Console.WriteLine("Workbook saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
