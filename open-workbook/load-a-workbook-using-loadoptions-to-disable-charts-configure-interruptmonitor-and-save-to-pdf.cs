using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create an interrupt monitor. The parameter indicates whether to terminate without throwing an exception.
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);
        // Start monitoring with a time limit (e.g., 5 seconds). Adjust as needed.
        monitor.StartMonitor(5000);

        // Configure load options and attach the interrupt monitor.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.InterruptMonitor = monitor;

        // If the library version supports it, disable chart loading to improve performance.
        // Uncomment the following line if the property exists:
        // loadOptions.DisableCharts = true;

        // Load the workbook using the constructor that accepts a file path and LoadOptions.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Attach the same interrupt monitor to the workbook for the save operation.
        workbook.InterruptMonitor = monitor;
        // Restart the monitor for the save phase.
        monitor.StartMonitor(5000);

        // Save the workbook as PDF using the Save method with a file name and format.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}