// Title: Cancel a long‑running Workbook.CalculateFormula call using a CancellationToken timeout in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a CancellationTokenSource, assigns a custom AbstractInterruptMonitor to a Workbook, and aborts CalculateFormula after a 2‑second timeout. | Show how to catch the CellsException with ExceptionType.Interrupted and then save the workbook containing partially calculated results. | Explain the steps to implement a CancellationInterruptMonitor class that checks CancellationToken.IsCancellationRequested during Aspose.Cells formula calculation.
// Common Searches: how to stop Aspose.Cells CalculateFormula after a certain time in C# | using CancellationToken to interrupt workbook formula evaluation in .NET | Aspose.Cells interrupt long running calculation without throwing unhandled exception | catch CellsException when calculation is cancelled in Aspose.Cells | save workbook after partial formula calculation when timeout occurs
// Tags: cancellation token interrupt monitor Aspose.Cells | timeout abort Workbook.CalculateFormula | handle CellsException Interrupted Aspose.Cells | partial workbook save after calculation cancellation | custom AbstractInterruptMonitor implementation C# | formula calculation timeout Aspose.Cells .NET

using System;
using System.Threading;
using Aspose.Cells;

// Custom interrupt monitor that checks a CancellationToken
// The example creates a workbook with many formulas, sets up a CancellationTokenSource that triggers after 2 seconds, and assigns a custom CancellationInterruptMonitor (derived from AbstractInterruptMonitor) to the workbook. When Workbook.CalculateFormula is called, the operation is interrupted if the timeout elapses, raising a CellsException with the Interrupted code. The code catches this exception, reports the interruption, and saves the workbook, which may contain partially calculated results.
class CancellationInterruptMonitor : AbstractInterruptMonitor
{
    private readonly CancellationToken _token;

    public CancellationInterruptMonitor(CancellationToken token)
    {
        _token = token;
    }

    // Return true when the token is cancelled
    public override bool IsInterruptionRequested => _token.IsCancellationRequested;

    // Let the library throw an exception when interrupted (default behavior)
    public override bool TerminateWithoutException => false;
}

class Program
{
    static void Main()
    {
        // Create a new workbook and add some data/formulas
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Add many formulas to make the calculation take noticeable time
        for (int i = 0; i < 50000; i++)
        {
            sheet.Cells[i, 3].Formula = $"=A1*A2+{i}";
        }

        // Set up a cancellation token that will be cancelled after a timeout (e.g., 2 seconds)
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            cts.CancelAfter(2000); // 2000 ms timeout

            // Assign the custom interrupt monitor to the workbook
            workbook.InterruptMonitor = new CancellationInterruptMonitor(cts.Token);

            try
            {
                // Perform formula calculation; it will be interrupted if the timeout elapses
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted due to timeout.");
            }
        }

        // Save the workbook (may contain partially calculated results)
        workbook.Save("InterruptedResult.xlsx");
    }
}
