using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

class CancelableCalculationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate it with sample data and formulas
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int i = 0; i < 20000; i++)
            {
                sheet.Cells[i, 0].PutValue(i);                     // Column A values
                sheet.Cells[i, 1].Formula = $"=A{i}+10";          // Column B formulas
            }

            // Set up a cancellation token source that the user can trigger
            CancellationTokenSource cts = new CancellationTokenSource();

            // Attach a custom interrupt monitor that checks the token
            workbook.InterruptMonitor = new TokenInterruptMonitor(cts.Token);

            // Optional: a calculation monitor to observe progress (can be omitted)
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = new SimpleCalcMonitor()
            };

            // Simulate a user abort after a short delay
            Task.Run(() =>
            {
                Thread.Sleep(500);               // Wait 0.5 seconds
                Console.WriteLine("User requested abort.");
                cts.Cancel();                    // Signal cancellation
            });

            Console.WriteLine("Starting formula calculation...");
            workbook.CalculateFormula(calcOptions);
            Console.WriteLine("Calculation finished normally.");

            // Save the workbook (partial results may be present)
            string outputPath = "CancelableResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            // Calculation was interrupted by the cancellation token
            Console.WriteLine("Calculation was interrupted by the cancellation token.");
        }
        catch (Exception ex)
        {
            // General exception handling
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Custom interrupt monitor that reads from a CancellationToken
class TokenInterruptMonitor : AbstractInterruptMonitor
{
    private readonly CancellationToken _token;

    public TokenInterruptMonitor(CancellationToken token)
    {
        _token = token;
    }

    // Return true when cancellation is requested
    public override bool IsInterruptionRequested => _token.IsCancellationRequested;
}

// Simple calculation monitor (optional, can be left empty)
class SimpleCalcMonitor : AbstractCalculationMonitor
{
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        // No action needed before each cell calculation
    }

    public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
    {
        // No action needed after each cell calculation
    }
}