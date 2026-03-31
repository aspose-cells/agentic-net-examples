using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Fill the first worksheet with a large amount of data to make the save operation lengthy
            Worksheet sheet = workbook.Worksheets[0];
            for (int i = 0; i < 20000; i++)
            {
                sheet.Cells[i, 0].Value = $"Row {i}";
            }

            // Create an InterruptMonitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            try
            {
                Console.WriteLine("Starting long save operation...");

                // In a separate thread, wait a short time then request interruption
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    Thread.Sleep(1000); // wait 1 second
                    Console.WriteLine("Requesting interruption...");
                    monitor.Interrupt(); // signal the monitor
                });

                // Attempt to save the workbook as PDF (this is the long operation)
                workbook.Save("InterruptedOutput.pdf", SaveFormat.Pdf);
                Console.WriteLine("Save completed (unexpected).");
            }
            catch (CellsException ex)
            {
                // Check if the exception was caused by an interruption request
                if (ex.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Operation was successfully interrupted.");
                }
                else
                {
                    Console.WriteLine($"Unexpected CellsException: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected exception: {ex.Message}");
            }
        }
    }
}