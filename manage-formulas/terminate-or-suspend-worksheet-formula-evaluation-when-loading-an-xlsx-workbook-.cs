using System;
using Aspose.Cells;

namespace AsposeCellsFormulaInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Prepare load options
            LoadOptions loadOptions = new LoadOptions();

            // Skip parsing formulas during load for performance or to suspend evaluation
            loadOptions.ParsingFormulaOnOpen = false;

            // Create an interrupt monitor that will terminate silently when interrupted
            // Setting terminateWithoutException to true means no exception will be thrown
            SystemTimeInterruptMonitor interruptMonitor = new SystemTimeInterruptMonitor(terminateWithoutException: true);

            // Assign the monitor to the load options
            loadOptions.InterruptMonitor = interruptMonitor;

            // Start monitoring with a short time limit (e.g., 200 ms)
            // If loading exceeds this time, the operation will be interrupted
            interruptMonitor.StartMonitor(200);

            Workbook workbook;

            try
            {
                // Load the workbook with the specified options
                workbook = new Workbook(sourcePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Loading was interrupted; create an empty workbook as a fallback
                Console.WriteLine("Loading was interrupted by the monitor.");
                workbook = new Workbook();
            }

            // At this point formulas have not been parsed.
            // If later you need to parse them, you can call ParseFormulas.
            // The 'ignoreError' flag determines whether invalid formulas throw exceptions.
            workbook.ParseFormulas(ignoreError: true);
            Console.WriteLine("Formulas parsed (if any).");

            // Save the workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}