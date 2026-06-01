using System;
using Aspose.Cells;

namespace AsposeCellsLightCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Create a SystemTimeInterruptMonitor with terminateWithoutException = false
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

            // Prepare load options and assign the interrupt monitor
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = monitor
            };

            // Start the monitor with a 30‑second (30000 ms) time limit
            monitor.StartMonitor(30000);

            try
            {
                // Load the workbook using the load options (LightCells API can be used for saving later)
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Save the loaded workbook to PDF format
                workbook.Save("output.pdf", SaveFormat.Pdf);

                Console.WriteLine("Workbook loaded and saved to PDF successfully.");
            }
            catch (Exception ex)
            {
                // If the operation exceeds the time limit, an interruption exception will be thrown
                Console.WriteLine($"Operation interrupted or failed: {ex.Message}");
            }
        }
    }
}