using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (replace with an actual file path)
            string sourcePath = "SampleWithFormulas.xlsx";

            // Create an interrupt monitor that can be triggered from another thread
            InterruptMonitor monitor = new InterruptMonitor();

            // Configure load options to use the interrupt monitor
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Simulate an external request to interrupt the operation after a short delay
            Task.Run(() =>
            {
                Thread.Sleep(500); // wait 0.5 seconds
                Console.WriteLine("Requesting interruption...");
                monitor.Interrupt(); // set IsInterruptionRequested to true
            });

            try
            {
                // Load the workbook with the interrupt monitor attached
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Force full formula calculation after loading
                // (CalculateOnOpen only affects saved files, so we call it explicitly)
                workbook.CalculateFormula();

                Console.WriteLine("Workbook loaded and calculated successfully.");
                
                // Save the workbook to verify that the operation completed (if not interrupted)
                workbook.Save("Result.xlsx");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // The operation was interrupted by the monitor
                Console.WriteLine("Formula calculation was interrupted successfully.");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}