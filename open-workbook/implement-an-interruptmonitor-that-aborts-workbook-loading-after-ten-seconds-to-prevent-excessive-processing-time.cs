using System;
using Aspose.Cells;

class InterruptLoadDemo
{
    static void Main()
    {
        // Path to the workbook that may take a long time to load
        string inputPath = "Large.xlsx";

        // Create a SystemTimeInterruptMonitor.
        // 'false' means an exception will be thrown when interruption occurs.
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Assign the monitor to LoadOptions.
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start monitoring with a 10‑second (10000 ms) time limit.
        monitor.StartMonitor(10000);

        try
        {
            // Load the workbook using the configured LoadOptions.
            Workbook wb = new Workbook(inputPath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Loading was interrupted after exceeding the time limit.
            Console.WriteLine("Loading aborted: " + ex.Message);
        }
    }
}