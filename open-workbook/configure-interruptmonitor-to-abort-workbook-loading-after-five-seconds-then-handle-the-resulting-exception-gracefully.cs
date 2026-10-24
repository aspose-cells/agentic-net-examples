// Title: Use Aspose.Cells InterruptMonitor in C# to abort loading a large workbook after 5 seconds and handle the timeout exception
// Common Searches: C# Aspose.Cells load large Excel file with a 5‑second timeout | How to stop workbook loading using InterruptMonitor in Aspose.Cells | Catch and process the 'Operation aborted' exception from Aspose.Cells | Implement a custom AbstractInterruptMonitor to limit Excel load time in .NET | Measure and abort Excel workbook loading time with Aspose.Cells LoadOptions
// Tags: Aspose.Cells time‑based interrupt monitor | C# workbook loading timeout using LoadOptions | custom interrupt monitor for Excel load | operation aborted exception handling Aspose.Cells | stop long workbook load Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example creates a TimeInterruptMonitor that checks a Stopwatch and requests interruption when more than five seconds have elapsed. The monitor is assigned to LoadOptions.InterruptMonitor, the stopwatch is started, and the workbook is loaded. If loading exceeds the time limit, Aspose.Cells throws an exception containing "Operation aborted", which is caught and handled gracefully while ensuring the stopwatch is stopped.
class Program
{
    static void Main()
    {
        // Path to the Excel file to load
        string filePath = "largeWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Stopwatch to track elapsed loading time
        Stopwatch stopwatch = new Stopwatch();

        // Create a custom interrupt monitor that aborts after 5 seconds
        var interruptMonitor = new TimeInterruptMonitor(stopwatch, 5);

        // Configure LoadOptions with the custom interrupt monitor
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = interruptMonitor
        };

        try
        {
            // Start timing before loading
            stopwatch.Start();

            // Load the workbook with the interrupt monitor in place
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Stop timing after successful load
            stopwatch.Stop();

            Console.WriteLine($"Workbook loaded successfully in {stopwatch.Elapsed.TotalSeconds:F2} seconds.");
            // Additional processing can be done here
        }
        catch (Exception ex)
        {
            // Ensure the stopwatch is stopped
            stopwatch.Stop();

            // Aspose.Cells throws a generic exception with message containing "Operation aborted" when interrupted
            if (ex.Message.Contains("Operation aborted"))
            {
                Console.WriteLine("Workbook loading was aborted after 5 seconds.");
            }
            else
            {
                Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
            }
        }
    }

    // Custom interrupt monitor that checks elapsed time against a threshold
    private class TimeInterruptMonitor : AbstractInterruptMonitor
    {
        private readonly Stopwatch _stopwatch;
        private readonly double _maxSeconds;

        public TimeInterruptMonitor(Stopwatch stopwatch, double maxSeconds)
        {
            _stopwatch = stopwatch;
            _maxSeconds = maxSeconds;
        }

        // Aspose.Cells checks this property to determine if interruption is requested
        public override bool IsInterruptionRequested => _stopwatch.Elapsed.TotalSeconds > _maxSeconds;
    }
}
