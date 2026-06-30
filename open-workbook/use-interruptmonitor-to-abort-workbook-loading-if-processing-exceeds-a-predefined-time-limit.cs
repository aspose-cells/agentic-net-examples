using System;
using Aspose.Cells;

public class InterruptMonitorDemo
{
    public static void Main()
    {
        // Author: Demo of ThreadInterruptMonitor to abort loading if it exceeds a time limit

        // Create a monitor; false => throws exception when interruption occurs
        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(false);

        // Attach the monitor to LoadOptions
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start monitoring with a 2‑second limit for the load operation
        monitor.StartMonitor(2000);

        Workbook workbook = null;
        try
        {
            // Load the workbook; will be interrupted if it takes longer than 2 seconds
            workbook = new Workbook("Large.xlsx", loadOptions);

            // Loading finished within the time limit
            monitor.FinishMonitor();

            // Optionally monitor the save operation with a new limit (1.5 seconds)
            monitor.StartMonitor(1500);
            workbook.Save("Result.xlsx");
            monitor.FinishMonitor();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Operation interrupted or failed: " + ex.Message);
        }
    }
}