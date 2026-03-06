using System;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsInterruptDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("First");
            sheet.Cells["A2"].PutValue("Second");
            sheet.Cells["A3"].Formula = "=A1+A2";

            // Create an interrupt monitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Start a background task that will request interruption after a short delay
            Task.Run(() =>
            {
                // Wait 500 ms before interrupting the save operation
                System.Threading.Thread.Sleep(500);
                Console.WriteLine("Requesting interruption...");
                monitor.Interrupt();
            });

            try
            {
                Console.WriteLine("Starting PDF conversion...");
                // Attempt to save the workbook as PDF; this should be interrupted
                workbook.Save("output.pdf", SaveFormat.Pdf);
                Console.WriteLine("PDF saved successfully (unexpected).");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Expected path when the operation is interrupted
                Console.WriteLine("PDF conversion was successfully interrupted.");
            }
            catch (Exception ex)
            {
                // Any other unexpected exceptions
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}