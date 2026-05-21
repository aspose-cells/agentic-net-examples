using System;
using Aspose.Cells;

class InterruptLoadDemo
{
    static void Main()
    {
        // Create an interrupt monitor that throws an exception when interrupted
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Assign the monitor to load options
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.InterruptMonitor = monitor;

        // Set a time limit (in milliseconds) for the loading operation
        int timeLimitMs = 2000; // 2 seconds
        monitor.StartMonitor(timeLimitMs);

        try
        {
            // Load the workbook with the specified load options
            Workbook wb = new Workbook("LargeFile.xlsx", loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle interruption or other loading errors
            Console.WriteLine("Loading was interrupted: " + ex.Message);
        }
    }
}