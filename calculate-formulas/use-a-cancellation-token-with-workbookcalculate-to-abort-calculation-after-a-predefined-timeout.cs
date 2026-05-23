using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsCancellationDemo
{
    // Custom interrupt monitor that checks a CancellationToken
    public class TokenInterruptMonitor : AbstractInterruptMonitor
    {
        private readonly CancellationToken _token;

        public TokenInterruptMonitor(CancellationToken token)
        {
            _token = token;
        }

        // Called by Aspose.Cells during long‑running operations
        public override bool IsInterruptionRequested => _token.IsCancellationRequested;

        // Keep default behavior (throw exception when interrupted)
        public override bool TerminateWithoutException => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a workbook and add many formulas to simulate a long calculation
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill column A with numbers
            for (int i = 0; i < 5000; i++)
            {
                cells[i, 0].PutValue(i + 1);
            }

            // Add a heavy formula that sums the whole column
            cells[0, 1].Formula = $"=SUM(A1:A{cells.MaxDataRow + 1})";

            // Set up a cancellation token that will be triggered after 2 seconds
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                cts.CancelAfter(TimeSpan.FromSeconds(2));

                // Assign the custom interrupt monitor to the workbook
                workbook.InterruptMonitor = new TokenInterruptMonitor(cts.Token);

                try
                {
                    // Start calculation; it will be aborted if the token is cancelled
                    workbook.CalculateFormula();
                    Console.WriteLine("Calculation completed successfully.");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Calculation was interrupted due to timeout.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error: {ex.Message}");
                }
            }

            // Save the workbook (if any calculations succeeded)
            workbook.Save("CancellationResult.xlsx");
        }
    }
}