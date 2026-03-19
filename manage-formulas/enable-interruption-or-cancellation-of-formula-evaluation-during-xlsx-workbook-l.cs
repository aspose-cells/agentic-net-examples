using System;
using Aspose.Cells;

class InterruptDuringLoadDemo
{
    static void Main()
    {
        // Create a time‑based interrupt monitor.
        // terminateWithoutException = false will cause an exception when interrupted.
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);

        // Start the monitor with a time limit (e.g., 500 milliseconds).
        // If the loading operation exceeds this limit, IsInterruptionRequested becomes true.
        monitor.StartMonitor(500);

        // Configure load options:
        // 1. Attach the interrupt monitor so the loading process can be cancelled.
        // 2. Disable formula parsing on open to avoid unnecessary calculations during load.
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor,
            ParsingFormulaOnOpen = false
        };

        try
        {
            // Load the workbook using the configured options.
            Workbook workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);

            Console.WriteLine("Workbook loaded successfully.");

            // Perform formula calculation only when needed.
            workbook.CalculateFormula();

            // Save the workbook (optional).
            workbook.Save("LoadedWorkbook.xlsx");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            // The loading operation was interrupted according to the monitor.
            Console.WriteLine("Loading was interrupted: " + ex.Message);
        }
    }
}