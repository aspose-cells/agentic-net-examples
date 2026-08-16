// Title: Cancel Aspose.Cells CalculateFormula with a CancellationToken after a timeout (C#)
// Description: Demonstrates how to abort a long‑running workbook.CalculateFormula call by assigning a custom InterruptMonitor that checks a CancellationToken. The token is set to cancel after a defined interval, causing Aspose.Cells to throw a CellsException with code Interrupted, which can be caught and handled.
// Keywords: Aspose.Cells | CalculateFormula | CancellationToken | InterruptMonitor | timeout cancellation | CellsException Interrupted | C# spreadsheet calculation | abort formula evaluation
// Common Searches: Aspose.Cells cancel CalculateFormula timeout | C# use CancellationToken with Aspose.Cells interrupt monitor | stop long running formula calculation Aspose.Cells | handle CellsException Interrupted | set calculation timeout Aspose.Cells workbook
// Developer Intent: The developer needs to stop a workbook.CalculateFormula operation if it exceeds a predefined time limit.
// Use Cases: Prevent UI freeze in a desktop app by cancelling heavy formula processing after a set duration. | Enforce server‑side execution limits for spreadsheet services to avoid runaway tasks. | Provide an API endpoint that lets callers abort calculation via a CancellationToken.
// AI Prompts: Create a reusable method that runs workbook.CalculateFormula with a configurable timeout and returns true if completed, false if cancelled. | Show logging of calculation start, end, and timeout events while using Aspose.Cells' interrupt monitor. | Write a unit test that verifies CalculateFormula is interrupted when the CancellationToken is cancelled after a short delay.

using System;
using System.Threading;
using Aspose.Cells;

// Demonstrates how to abort a long‑running workbook.CalculateFormula call by assigning a custom InterruptMonitor that checks a CancellationToken. The token is set to cancel after a defined interval, causing Aspose.Cells to throw a CellsException with code Interrupted, which can be caught and handled.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data with formulas
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate many rows to make calculation take noticeable time
        for (int i = 0; i < 2000; i++)
        {
            sheet.Cells[i, 0].PutValue(i);
            sheet.Cells[i, 1].Formula = $"=A{i}+SUM(A1:A{i})";
        }

        // Set up a cancellation token that will be triggered after a timeout
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Define timeout (e.g., 1500 milliseconds)
            cts.CancelAfter(1500);

            // Assign a custom interrupt monitor that checks the token
            workbook.InterruptMonitor = new CancellationInterruptMonitor(cts.Token);

            try
            {
                // Perform formula calculation; it will be interrupted if the token is cancelled
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted due to timeout.");
            }
        }

        // Save the workbook (optional)
        workbook.Save("Result.xlsx");
    }

    // Custom interrupt monitor that uses a CancellationToken to request interruption
    private class CancellationInterruptMonitor : AbstractInterruptMonitor
    {
        private readonly CancellationToken _token;

        public CancellationInterruptMonitor(CancellationToken token)
        {
            _token = token;
        }

        // Return true when the token signals cancellation
        public override bool IsInterruptionRequested => _token.IsCancellationRequested;

        // Keep default behavior: throw CellsException when interrupted
        public override bool TerminateWithoutException => false;
    }
}
