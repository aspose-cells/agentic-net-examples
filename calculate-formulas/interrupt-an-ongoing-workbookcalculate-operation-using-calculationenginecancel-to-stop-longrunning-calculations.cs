// Title: Cancel a long‑running workbook.CalculateFormula with Aspose.Cells InterruptMonitor (C#)
// Description: Demonstrates how to assign an InterruptMonitor to a workbook, trigger an interrupt from a background task, and catch the Interrupted exception to stop a time‑consuming workbook.CalculateFormula operation in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | InterruptMonitor | Cancel workbook.CalculateFormula | CalculationEngine.Cancel | C# | .NET | long running formula calculation | Interrupted exception | Excel calculation cancellation | performance optimization
// Common Searches: how to stop workbook.CalculateFormula in Aspose.Cells | Aspose.Cells interrupt long calculation C# example | cancel Excel formula evaluation with Aspose.Cells | InterruptMonitor usage Aspose.Cells .NET | handle CellsException Interrupted Aspose.Cells
// Developer Intent: Terminate an ongoing workbook.CalculateFormula call by signaling an interrupt.
// Use Cases: Prevent UI freeze by aborting heavy spreadsheet recalculations after a timeout. | Provide a cancel button that stops formula evaluation in desktop or web apps. | Save a workbook with partially calculated data when the calculation is interrupted.
// AI Prompts: Generate C# code that uses Aspose.Cells InterruptMonitor to cancel workbook.CalculateFormula after 1 second and handle the Interrupted exception. | Explain how to integrate calculation interruption into an ASP.NET MVC app with Aspose.Cells, including user feedback for cancellation. | Write unit tests that verify workbook.CalculateFormula is interrupted when InterruptMonitor.Interrupt is invoked.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    // Demonstrates how to assign an InterruptMonitor to a workbook, trigger an interrupt from a background task, and catch the Interrupted exception to stop a time‑consuming workbook.CalculateFormula operation in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large number of cells with formulas to make calculation time‑consuming
            for (int row = 0; row < 20000; row++)
            {
                // Simple formula that depends on the previous row to create a chain of calculations
                if (row == 0)
                    sheet.Cells[row, 0].Formula = "=1";
                else
                    sheet.Cells[row, 0].Formula = $"=A{row}+1";
            }

            // Create an interrupt monitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Start a background task that will request interruption after a short delay
            Task.Run(() =>
            {
                Thread.Sleep(500); // Wait 0.5 seconds before interrupting
                Console.WriteLine("Requesting interruption...");
                monitor.Interrupt(); // Signal the interrupt
            });

            // Attempt to calculate all formulas; this operation should be interrupted
            try
            {
                Console.WriteLine("Starting calculation...");
                workbook.CalculateFormula(); // Long‑running operation
                Console.WriteLine("Calculation completed without interruption (unexpected).");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Expected path when the operation is interrupted
                Console.WriteLine("Calculation was successfully interrupted.");
            }
            catch (Exception ex)
            {
                // Any other unexpected exception
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            // Save the workbook (partial results may be present)
            try
            {
                workbook.Save("InterruptedResult.xlsx");
                Console.WriteLine("Workbook saved (partial data may be present).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
