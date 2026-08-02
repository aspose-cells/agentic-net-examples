// Title: Cancel Aspose.Cells CalculateFormula with a CancellationToken after a timeout (C#)
// Description: Demonstrates how to abort a long‑running workbook.CalculateFormula call by attaching a custom InterruptMonitor that checks a CancellationToken. The token is set to fire after a predefined interval, the calculation is stopped, a CellsException with the Interrupted code is caught, and the partially processed workbook can be saved.
// Keywords: Aspose.Cells | CalculateFormula | CancellationToken | InterruptMonitor | C# | .NET | calculation timeout | abort formula evaluation | CellsException | Interrupted | long running formulas
// Common Searches: Aspose.Cells stop CalculateFormula after X seconds | use CancellationToken to interrupt Aspose.Cells calculation | C# interrupt monitor for workbook.CalculateFormula | timeout for long formula evaluation Aspose.Cells | handle CellsException Interrupted in .NET
// Developer Intent: The developer needs to terminate a lengthy CalculateFormula operation automatically when a time limit is reached.
// Use Cases: Process large spreadsheets in a web service but enforce a maximum execution time. | Prevent server‑side hangs caused by complex or malicious formulas. | Save a workbook with whatever results were computed before the timeout for later analysis.
// AI Prompts: Write C# code that creates a CancellationTokenSource with a 5‑second timeout, implements a custom InterruptMonitor, attaches it to an Aspose.Cells workbook, and calls CalculateFormula. | Explain how to catch the CellsException for an interrupted calculation and log the elapsed time before the cancellation. | Show how to modify the InterruptMonitor so that it silently stops calculation without throwing an exception.

using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Custom interrupt monitor that checks a CancellationToken
    // Demonstrates how to abort a long‑running workbook.CalculateFormula call by attaching a custom InterruptMonitor that checks a CancellationToken. The token is set to fire after a predefined interval, the calculation is stopped, a CellsException with the Interrupted code is caught, and the partially processed workbook can be saved.
    public class CancellationInterruptMonitor : AbstractInterruptMonitor
    {
        private readonly CancellationToken _token;

        public CancellationInterruptMonitor(CancellationToken token)
        {
            _token = token;
        }

        // Called by Aspose.Cells during long‑running operations
        public override bool IsInterruptionRequested => _token.IsCancellationRequested;

        // Keep default behavior (throw CellsException when interrupted)
        public override bool TerminateWithoutException => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a cancellation token that will be triggered after 3 seconds
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3));

            // Create a new workbook and fill it with many formulas to simulate a long calculation
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate 10,000 rows with simple formulas (adjust count for longer processing if needed)
            const int rowCount = 10000;
            for (int i = 0; i < rowCount; i++)
            {
                cells[i, 0].PutValue(i + 1);                     // Column A: numbers
                cells[i, 1].Formula = $"=A{i + 1}*2";           // Column B: formula depending on column A
            }

            // Assign the custom interrupt monitor to the workbook
            workbook.InterruptMonitor = new CancellationInterruptMonitor(cts.Token);

            try
            {
                // Start formula calculation; it will be aborted when the token is cancelled
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted due to timeout.");
            }

            // Save the workbook only if needed (will contain partially calculated results)
            workbook.Save("InterruptedResult.xlsx");
        }
    }
}
