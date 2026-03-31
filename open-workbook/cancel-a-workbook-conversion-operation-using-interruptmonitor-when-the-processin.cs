using System;
using Aspose.Cells;

class WorkbookConversionWithInterrupt
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        for (int i = 0; i < 5000; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i}");
        }

        // Create a ThreadInterruptMonitor (throws exception on interruption)
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(false);
        // Assign the monitor to the workbook
        workbook.InterruptMonitor = monitor;

        // Start monitoring with a 1‑second time limit
        monitor.StartMonitor(1000);

        try
        {
            // Attempt to convert the workbook to PDF (this operation will be monitored)
            workbook.Save("Converted.pdf", SaveFormat.Pdf);
            Console.WriteLine("Conversion completed successfully.");
        }
        catch (CellsException ex)
        {
            // Check if the exception was caused by an interruption
            if (ex.Code == ExceptionType.Interrupted)
                Console.WriteLine("Conversion was interrupted because the time limit was exceeded.");
            else
                Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Finish the monitor to release resources
            monitor.FinishMonitor();
        }
    }
}