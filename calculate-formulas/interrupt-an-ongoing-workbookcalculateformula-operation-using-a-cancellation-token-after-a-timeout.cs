using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// Custom interrupt monitor that checks a CancellationToken
public class TokenInterruptMonitor : AbstractInterruptMonitor
{
    private readonly CancellationToken _token;

    public TokenInterruptMonitor(CancellationToken token)
    {
        _token = token;
    }

    // Return true when the token signals cancellation
    public override bool IsInterruptionRequested => _token.IsCancellationRequested;

    // Let Aspose.Cells throw an exception when interrupted (default behavior)
    public override bool TerminateWithoutException => false;
}

public class CalculateFormulaWithTimeoutDemo
{
    public static void Run()
    {
        // Create a cancellation token source that will cancel after 2 seconds
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));

        // Assign the custom monitor to the workbook
        Workbook workbook = new Workbook();
        workbook.InterruptMonitor = new TokenInterruptMonitor(cts.Token);

        // Populate some data and formulas to make calculation take noticeable time
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill a large range with values
        for (int row = 0; row < 5000; row++)
        {
            cells[row, 0].PutValue(row);
        }

        // Add a formula that sums a large range (will be time‑consuming)
        cells["B1"].Formula = $"=SUM(A1:A{5000})";

        try
        {
            // Start calculation; it will be interrupted if the token fires
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

        // Save the workbook (optional, will succeed even if calculation was interrupted)
        try
        {
            workbook.Save("CalculateFormulaWithTimeoutDemo.xlsx");
            Console.WriteLine("Workbook saved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}

// Entry point
class Program
{
    static void Main()
    {
        CalculateFormulaWithTimeoutDemo.Run();
    }
}