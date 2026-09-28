// Title: How to interrupt a long‑running workbook.CalculateFormula call using Aspose.Cells InterruptMonitor in C#
// AI Prompts: Generate C# code that configures an Aspose.Cells InterruptMonitor to cancel a workbook.CalculateFormula execution after a timed delay. | Show how to catch CellsException with ExceptionType.Interrupted during both formula calculation and workbook saving in Aspose.Cells. | Explain how to launch a background task that triggers CalculationEngine.Cancel while formulas are being evaluated with Aspose.Cells.
// Common Searches: c# aspnet interrupt workbook.CalculateFormula after 1 second using Aspose.Cells | Aspose.Cells how to cancel long running formula calculation in .NET | example of using InterruptMonitor to stop Excel calculation with Aspose.Cells | handling interrupted calculation exception when saving workbook in C# Aspose.Cells | stop calculation engine in Aspose.Cells without freezing UI
// Tags: Aspose.Cells InterruptMonitor for calculation cancellation | cancel workbook.CalculateFormula in C# | handle CellsException Interrupted in Aspose.Cells | background task to stop Excel formula evaluation | saving workbook after interrupted calculation Aspose.Cells

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // The sample creates a workbook with 20,000 rows of growing SUM formulas, assigns an InterruptMonitor to the workbook, and starts a background task that calls monitor.Interrupt() after 500 ms. The workbook.CalculateFormula() call runs inside a try‑catch that specifically handles CellsException with ExceptionType.Interrupted, and the save operation is wrapped similarly to manage possible interruption during file writing.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large range with formulas to make calculation time‑consuming
            // Example: each cell in column B sums a range in column A
            for (int row = 0; row < 20000; row++)
            {
                cells[row, 0].PutValue(row + 1);                     // Column A: simple numbers
                cells[row, 1].Formula = $"=SUM(A1:A{row + 1})";      // Column B: growing SUM formula
            }

            // Create an interrupt monitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Start a background task that will request interruption after a short delay
            Task.Run(() =>
            {
                Thread.Sleep(500); // Wait 0.5 seconds before interrupting
                Console.WriteLine("Requesting interruption...");
                monitor.Interrupt(); // Signal the calculation to stop
            });

            // Perform the long‑running calculation and handle possible interruption
            try
            {
                Console.WriteLine("Starting calculation...");
                workbook.CalculateFormula(); // This may be interrupted by the monitor
                Console.WriteLine("Calculation completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted as requested.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error during calculation: {ex.Message}");
            }

            // Attempt to save the workbook; handle possible interruption exception during save
            try
            {
                workbook.Save("InterruptedResult.xlsx");
                Console.WriteLine("Workbook saved.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Saving may also be interrupted; still attempt to write the file if possible
                Console.WriteLine("Save operation was interrupted, but partial results may have been written.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error during save: {ex.Message}");
            }
        }
    }
}
