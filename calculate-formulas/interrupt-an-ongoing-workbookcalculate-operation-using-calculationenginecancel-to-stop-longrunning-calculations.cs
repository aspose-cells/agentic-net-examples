// Title: Cancel a long‑running CalculateFormula in Aspose.Cells using InterruptMonitor (C#)
// Description: Shows how to attach an InterruptMonitor to a workbook, trigger interruption from a background task, run Workbook.CalculateFormula, catch the Interrupted CellsException, and optionally save the partially calculated file.
// Keywords: Aspose.Cells | InterruptMonitor | Cancel CalculateFormula | C# workbook calculation | CellsException Interrupted | long running formulas | stop calculation Aspose | aspose cells interrupt calculation | aspose cells .NET
// Common Searches: aspnet cancel CalculateFormula | interrupt Aspose.Cells calculation | how to stop workbook.CalculateFormula C# | use InterruptMonitor Aspose.Cells | catch CellsException Interrupted
// Developer Intent: Need to abort a workbook.CalculateFormula that exceeds a time limit.
// Use Cases: Abort massive formula evaluation when a user clicks Cancel in a desktop app. | Prevent server time‑outs by terminating calculations that run longer than allowed. | Obtain partial results for diagnostics after an interrupted calculation. | Implement a timeout‑based safety net for automated spreadsheet processing.
// AI Prompts: Provide C# code that sets up an InterruptMonitor to stop CalculateFormula after 1 second. | Show how to catch the CellsException with ExceptionType.Interrupted and log the interruption. | Explain best practices for saving a workbook after an interrupted calculation to avoid corruption. | Generate a unit test that verifies CalculateFormula can be cancelled using InterruptMonitor.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Shows how to attach an InterruptMonitor to a workbook, trigger interruption from a background task, run Workbook.CalculateFormula, catch the Interrupted CellsException, and optionally save the partially calculated file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large number of formulas to make calculation take noticeable time
            for (int row = 0; row < 20000; row++)
            {
                // Simple formula referencing previous row to create dependency chain
                if (row == 0)
                    sheet.Cells[row, 0].PutValue(1);
                else
                    sheet.Cells[row, 0].Formula = $"=A{row}+1";
            }

            // Create an interrupt monitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Start a background task that will request interruption after a short delay
            Task.Run(() =>
            {
                Thread.Sleep(500); // wait 0.5 seconds
                Console.WriteLine("Requesting interruption...");
                monitor.Interrupt(); // trigger interruption
            });

            try
            {
                Console.WriteLine("Starting long-running calculation...");
                // Perform calculation (lifecycle rule: calculate)
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed without interruption.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was interrupted successfully.");
            }

            // Attempt to save the workbook (lifecycle rule: save)
            try
            {
                workbook.Save("InterruptedResult.xlsx");
                Console.WriteLine("Workbook saved (may contain partial results).");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error during save: {saveEx.Message}");
            }
        }
    }
}
