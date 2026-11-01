// Title: How to abort loading a large Excel workbook after 10 seconds with Aspose.Cells InterruptMonitor in C#
// Common Searches: c# Aspose.Cells load large xlsx with timeout using InterruptMonitor | stop workbook loading after a specific time limit in Aspose.Cells | example of TimeBasedInterruptMonitor to cancel Excel file load in .NET | prevent long‑running workbook load in Aspose.Cells by setting a time limit | use LoadOptions.InterruptMonitor to abort Excel loading after 10 seconds
// Tags: Aspose.Cells InterruptMonitor timeout | C# LoadOptions time limit for XLSX | abort large workbook load with time based monitor | Excel file loading cancellation using Aspose.Cells | custom InterruptMonitor implementation C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Custom interrupt monitor that aborts after a specified time span
// The sample defines a TimeBasedInterruptMonitor that inherits from InterruptMonitor and returns true after ten seconds. The monitor is assigned to LoadOptions.InterruptMonitor, so loading a large XLSX workbook with Aspose.Cells is automatically aborted when the time limit is exceeded, and the exception is handled gracefully.
public class TimeBasedInterruptMonitor : InterruptMonitor
{
    private readonly Stopwatch _stopwatch;
    private readonly TimeSpan _maxDuration;

    public TimeBasedInterruptMonitor(TimeSpan maxDuration)
    {
        _maxDuration = maxDuration;
        _stopwatch = Stopwatch.StartNew();
    }

    // Implement the interruption logic required by InterruptMonitor
    public bool IsInterrupted()
    {
        return _stopwatch.Elapsed > _maxDuration;
    }
}

public class WorkbookLoader
{
    public static void Main()
    {
        string inputPath = "largeWorkbook.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Create an interrupt monitor that stops loading after 10 seconds
        var monitor = new TimeBasedInterruptMonitor(TimeSpan.FromSeconds(10));

        // Set up load options and attach the monitor
        var loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            InterruptMonitor = monitor
        };

        try
        {
            // Load the workbook with the interrupt monitor active
            var workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform additional processing here

            // Save the workbook if needed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Loading was aborted due to timeout or another error
            Console.WriteLine("Workbook loading was aborted: " + ex.Message);
        }
    }
}
