using System;
using System.Threading;
using Aspose.Cells;

namespace InterruptWorkbookToPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large amount of data to make the PDF conversion take noticeable time
            for (int row = 0; row < 20000; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row}");
                sheet.Cells[row, 1].PutValue(row);
            }

            // Create an InterruptMonitor and assign it to the workbook
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Start a background task that will request interruption after a short delay
            ThreadPool.QueueUserWorkItem(_ =>
            {
                Thread.Sleep(1000); // wait 1 second
                Console.WriteLine("Requesting interruption...");
                monitor.Interrupt(); // signal interruption
            });

            try
            {
                Console.WriteLine("Starting PDF conversion...");
                // Attempt to save the workbook as PDF; this operation may be interrupted
                workbook.Save("output.pdf", SaveFormat.Pdf);
                Console.WriteLine("PDF conversion completed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                // Expected path when the operation is interrupted
                Console.WriteLine("PDF conversion was interrupted as requested.");
            }
            catch (Exception ex)
            {
                // Any other unexpected errors
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}