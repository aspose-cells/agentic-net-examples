using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class InterruptMonitorDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Create an interrupt monitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            try
            {
                Console.WriteLine("Starting save operation...");

                // Start a background task that will interrupt after 1 second
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    Thread.Sleep(1000); // wait 1 second
                    Console.WriteLine("Interrupting operation...");
                    monitor.Interrupt(); // request interruption
                });

                // Perform a potentially long-running conversion (save to PDF)
                workbook.Save("output.pdf", SaveFormat.Pdf);
                Console.WriteLine("Save completed successfully.");
            }
            catch (CellsException ex)
            {
                // Check if the exception was caused by interruption
                if (ex.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Operation was successfully interrupted.");
                }
                else
                {
                    Console.WriteLine("Error occurred: " + ex.Message);
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            InterruptMonitorDemo.Run();
        }
    }
}