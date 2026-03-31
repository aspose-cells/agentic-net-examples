using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the large source workbook and the PDF output
            string sourcePath = "LargeWorkbook.xlsx";
            string pdfPath = "LargeWorkbook.pdf";

            // Create an interrupt monitor that will throw an exception when interrupted
            // (terminateWithoutException = false)
            SystemTimeInterruptMonitor interruptMonitor = new SystemTimeInterruptMonitor(false);

            // Prepare load options and attach the interrupt monitor
            LoadOptions loadOptions = new LoadOptions
            {
                InterruptMonitor = interruptMonitor
            };

            // Prepare PDF save options (optional: ignore rendering errors)
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                // If you want the conversion to continue despite rendering errors
                // set IgnoreError = true;
                IgnoreError = true
            };

            // Define a maximum time (in milliseconds) for each operation segment.
            // Adjust these values based on the expected processing time.
            const int loadTimeLimitMs = 5000;   // 5 seconds for loading
            const int saveTimeLimitMs = 10000;  // 10 seconds for saving

            bool conversionCompleted = false;

            while (!conversionCompleted)
            {
                try
                {
                    // Start monitoring the load operation
                    interruptMonitor.StartMonitor(loadTimeLimitMs);

                    // Perform the conversion using the utility method that respects load options
                    // This will load the workbook, apply the interrupt monitor, and then save as PDF.
                    ConversionUtility.Convert(sourcePath, loadOptions, pdfPath, pdfSaveOptions);

                    // If we reach this point, conversion succeeded
                    conversionCompleted = true;
                    Console.WriteLine("Workbook successfully converted to PDF.");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
                {
                    // The operation was interrupted (time limit exceeded). Retry the whole process.
                    Console.WriteLine("Operation was interrupted. Retrying...");
                }
                catch (Exception ex)
                {
                    // Any other exception should be reported and the loop terminated.
                    Console.WriteLine($"Unexpected error: {ex.Message}");
                    break;
                }
                finally
                {
                    // Ensure the monitor is stopped before the next iteration.
                    // The monitor does not have an explicit Stop method; starting a new monitor
                    // in the next loop iteration will reset the timing.
                }
            }
        }
    }
}