// Title: Cancel Aspose.Cells Formula Calculation with a CancellationToken via InterruptMonitor (C#)
// Description: Shows how to abort a long‑running workbook.CalculateFormula() by wiring a CancellationToken to a custom InterruptMonitor, catching the Interrupted CellsException, and saving the partially calculated workbook.
// Keywords: Aspose.Cells | CancellationToken | InterruptMonitor | Abort formula calculation | C# | Workbook.CalculateFormula cancellation | Custom AbstractInterruptMonitor | CellsException Interrupted
// Common Searches: how to cancel Aspose.Cells CalculateFormula | Aspose.Cells InterruptMonitor example C# | cancel long running Excel formula calculation .NET | use CancellationToken with Aspose.Cells | stop workbook.CalculateFormula on user abort
// Developer Intent: Implement a cancellation mechanism that stops Aspose.Cells formula calculation when a user‑initiated CancellationToken is triggered.
// Use Cases: Provide a Cancel button in a WinForms/WPF app that aborts Excel calculations. | Enforce a maximum calculation time for large worksheets by timing out. | Allow ASP.NET Core endpoints to terminate formula evaluation if the request is cancelled.
// AI Prompts: Generate a timeout‑based CancellationTokenSource example for Aspose.Cells formula calculation. | Show code for handling cancellation in an ASP.NET Core controller that processes uploaded Excel files with Aspose.Cells. | Explain how to log the last successfully calculated cell range before an interruption occurs.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsCancellationDemo
{
    // Custom interrupt monitor that checks a CancellationToken.
    // Shows how to abort a long‑running workbook.CalculateFormula() by wiring a CancellationToken to a custom InterruptMonitor, catching the Interrupted CellsException, and saving the partially calculated workbook.
    public class CancellationInterruptMonitor : AbstractInterruptMonitor
    {
        private readonly CancellationToken _token;

        public CancellationInterruptMonitor(CancellationToken token)
        {
            _token = token;
        }

        // Returns true when cancellation is requested, causing Aspose.Cells to interrupt the operation.
        public override bool IsInterruptionRequested => _token.IsCancellationRequested;

        // Keep default behavior: throw CellsException when interrupted.
        public override bool TerminateWithoutException => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a cancellation source that will be triggered after a short delay.
            var cts = new CancellationTokenSource();

            // Simulate user abort after 1 second.
            Task.Run(() =>
            {
                Thread.Sleep(1000);
                Console.WriteLine("Cancellation requested by user.");
                cts.Cancel();
            });

            // Create a new workbook and populate it with sample data and formulas.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int i = 0; i < 5000; i++)
            {
                // Simple data to make calculation take some time.
                sheet.Cells[i, 0].PutValue(i);
                sheet.Cells[i, 1].Formula = $"=A{i}+B{i}";
            }

            // Assign the custom interrupt monitor to the workbook.
            workbook.InterruptMonitor = new CancellationInterruptMonitor(cts.Token);

            try
            {
                Console.WriteLine("Starting formula calculation...");
                // Perform calculation; it will be interrupted when the token is cancelled.
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed without interruption.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted as requested.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook (will save whatever has been calculated so far).
            workbook.Save("CancellationDemo.xlsx");
            Console.WriteLine("Workbook saved.");
        }
    }
}
