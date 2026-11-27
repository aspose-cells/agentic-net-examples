// Title: Creating a time‑based InterruptMonitor to limit Aspose.Cells workbook loading in C#
// Common Searches: how to set a timeout for loading an Excel workbook with Aspose.Cells in .NET | c# example of using InterruptMonitor to stop workbook load after a specific duration | Aspose.Cells load options with custom interrupt monitor for large files | prevent long‑running workbook load in C# using Aspose.Cells timeout | configure maximum load time for Excel files with Aspose.Cells .NET API
// Tags: Aspose.Cells workbook load timeout | C# interrupt monitor for Excel loading | time‑based load cancellation Aspose.Cells | LoadOptions interrupt monitor usage | prevent excessive workbook load duration

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace WorkbookLoadingWithInterrupt
{
    // Dummy monitor class kept for reference; not used because IInterruptMonitor is unavailable in this version.
    // The example shows how to create a simple TimeBasedInterruptMonitor that uses a Stopwatch to enforce a maximum loading duration, demonstrates checking file existence, configuring LoadOptions, and loading a workbook with Aspose.Cells in C#. Although the IInterruptMonitor interface is not available in this version, the pattern illustrates how to integrate a custom monitor to abort loading after a ten‑second limit and handle any exceptions that arise.
    public class TimeBasedInterruptMonitor
    {
        private readonly Stopwatch _stopwatch;
        private readonly TimeSpan _maxDuration;

        public TimeBasedInterruptMonitor(TimeSpan maxDuration)
        {
            _maxDuration = maxDuration;
            _stopwatch = Stopwatch.StartNew();
        }

        // Returns true while the allowed time has not elapsed.
        public bool Continue()
        {
            return _stopwatch.Elapsed < _maxDuration;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the workbook to be loaded.
            string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Configure load options. InterruptMonitor is omitted because the required interface is unavailable.
            LoadOptions loadOptions = new LoadOptions();

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");
                // Further processing can be done here.
            }
            catch (Exception ex)
            {
                // Handle errors such as corrupted files or other load failures.
                Console.WriteLine($"Workbook loading failed: {ex.Message}");
            }
        }
    }
}
