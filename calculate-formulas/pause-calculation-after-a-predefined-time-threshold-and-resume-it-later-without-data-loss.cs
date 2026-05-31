using System;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Load a workbook (replace with your actual file)
            LoadOptions loadOptions = new LoadOptions();

            // Create a SystemTimeInterruptMonitor that throws an exception when time limit is exceeded
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);
            loadOptions.InterruptMonitor = monitor;

            // Load the workbook with the interrupt monitor attached
            Workbook wb = new Workbook("Large.xlsx", loadOptions);

            // Assign the same monitor to the workbook (required for calculation)
            wb.InterruptMonitor = monitor;

            // First calculation attempt with a short time limit (e.g., 1 second)
            monitor.StartMonitor(1000); // 1000 ms = 1 second
            try
            {
                wb.CalculateFormula();
                Console.WriteLine("Calculation completed within the first time limit.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation paused after time threshold was reached.");

                // Resume calculation with a new (longer) time limit
                monitor.StartMonitor(3000); // 3 seconds for the remaining work
                try
                {
                    wb.CalculateFormula(); // Re‑run calculation; Aspose.Cells will continue where it left off
                    Console.WriteLine("Calculation resumed and completed successfully.");
                }
                catch (CellsException ex2) when (ex2.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Calculation was interrupted again. You can repeat the resume logic as needed.");
                    // Optionally handle further retries here
                }
            }

            // Save the workbook after successful calculation
            wb.Save("Result.xlsx");
            Console.WriteLine("Workbook saved.");
        }
    }
}