using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large number of cells with formulas to make calculation take time
            // Example: each cell in column A contains a formula that sums a growing range
            for (int i = 0; i < 20000; i++)
            {
                // Simple value cell
                sheet.Cells[i, 0].PutValue(i);
                // Formula that depends on the whole column up to the current row
                sheet.Cells[i, 1].Formula = $"=SUM(A1:A{i + 1})";
            }

            // Create an interrupt monitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Start a background task that will request interruption after a short delay
            Task.Run(() =>
            {
                Thread.Sleep(500); // wait 0.5 seconds
                Console.WriteLine("Requesting interruption...");
                monitor.Interrupt(); // Signal the interruption
            });

            try
            {
                Console.WriteLine("Starting long-running calculation...");
                // Perform calculation; this will be interrupted by the monitor
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed without interruption.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Expected path when interruption occurs
                Console.WriteLine("Calculation was successfully interrupted.");
            }
            catch (Exception ex)
            {
                // Any other unexpected exception
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook (optional, will contain partially calculated data)
            try
            {
                workbook.Save("InterruptedResult.xlsx");
                Console.WriteLine("Workbook saved (may contain partial results).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}