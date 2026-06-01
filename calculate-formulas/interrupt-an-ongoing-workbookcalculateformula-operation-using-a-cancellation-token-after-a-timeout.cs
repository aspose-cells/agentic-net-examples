using System;
using System.Threading;
using Aspose.Cells;

class CancellationInterruptMonitor : AbstractInterruptMonitor
{
    private readonly CancellationToken _token;

    public CancellationInterruptMonitor(CancellationToken token)
    {
        _token = token;
    }

    // Return true when the cancellation token is signaled
    public override bool IsInterruptionRequested => _token.IsCancellationRequested;

    // Throw CellsException when interrupted (default behavior)
    public override bool TerminateWithoutException => false;
}

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data with formulas
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        for (int i = 0; i < 20000; i++)
        {
            sheet.Cells[i, 0].PutValue(i);                     // Simple value
            sheet.Cells[i, 1].Formula = $"=A{i}+10";          // Formula that depends on the value
        }

        // Set up a cancellation token that will be triggered after 2 seconds
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            cts.CancelAfter(TimeSpan.FromSeconds(2));

            // Assign the custom interrupt monitor to the workbook
            workbook.InterruptMonitor = new CancellationInterruptMonitor(cts.Token);

            try
            {
                // Start formula calculation; it will be interrupted if the timeout expires
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed without interruption.");
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