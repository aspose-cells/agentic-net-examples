// Title: Cancel Aspose.Cells formula calculation using a CancellationToken in C#
// AI Prompts: Write C# code that creates a CancellationTokenSource, implements a custom AbstractCalculationMonitor that checks the token before each cell, and passes the monitor to Workbook.CalculateFormula to enable user‑initiated abort. | Show how to launch a background task that calls CancellationTokenSource.Cancel after a short delay, then catch the OperationCanceledException thrown by CalculateFormula. | Describe the steps to save the workbook after a cancelled calculation, preserving any cells that were already evaluated.
// Common Searches: asp.net cancel workbook.CalculateFormula after user request | c# use CancellationToken with Aspose.Cells calculation monitor | stop long running formula evaluation in Aspose.Cells using token | handle OperationCanceledException when calculating formulas in Aspose.Cells | save partially calculated workbook after cancellation Aspose.Cells
// Tags: cancellation token Aspose.Cells calculation monitor | custom AbstractCalculationMonitor for formula abort | Workbook.CalculateFormula cancellation handling | partial workbook save after calculation abort | OperationCanceledException Aspose.Cells formula evaluation

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace CancelCalculationDemoApp
{
    // The example creates a workbook with 10,000 rows of data and formulas, sets up a CancellationTokenSource, and defines a custom AbstractCalculationMonitor that throws OperationCanceledException when the token is signaled. Workbook.CalculateFormula is executed with this monitor, a background task cancels the token after 0.5 seconds, the cancellation is caught, and the workbook (containing any partially calculated results) is saved.
    class CancelCalculationDemo
    {
        static void Main()
        {
            try
            {
                // Create a workbook and populate it with sample data and formulas
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                for (int i = 0; i < 10000; i++)
                {
                    sheet.Cells[i, 0].PutValue(i);                     // Column A values
                    sheet.Cells[i, 1].Formula = $"=A{i}+10";          // Column B formulas
                }

                // Set up a cancellation token source that can be triggered by the user
                using CancellationTokenSource cts = new CancellationTokenSource();

                // Create a custom calculation monitor that checks the token before each cell calculation
                var calcMonitor = new CancellationCalculationMonitor(cts.Token);
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = calcMonitor
                };

                // Simulate a user abort after a short delay
                Task.Run(() =>
                {
                    Thread.Sleep(500); // Wait 0.5 seconds
                    Console.WriteLine("User requested cancellation.");
                    cts.Cancel();
                });

                try
                {
                    // Perform calculation with the monitor attached
                    workbook.CalculateFormula(options);
                    Console.WriteLine("Calculation completed successfully.");
                }
                catch (OperationCanceledException)
                {
                    // Expected path when cancellation is requested
                    Console.WriteLine("Calculation was aborted by the user.");
                }

                // Save the workbook (optional, will contain partially calculated results)
                string outputPath = "CancelledCalculation.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Custom monitor that aborts calculation when the cancellation token is set
        class CancellationCalculationMonitor : AbstractCalculationMonitor
        {
            private readonly CancellationToken _token;

            public CancellationCalculationMonitor(CancellationToken token)
            {
                _token = token;
            }

            public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
            {
                if (_token.IsCancellationRequested)
                {
                    // Throw an operation cancelled exception to stop the calculation engine
                    throw new OperationCanceledException("Calculation cancelled via token.");
                }
            }
        }
    }
}
