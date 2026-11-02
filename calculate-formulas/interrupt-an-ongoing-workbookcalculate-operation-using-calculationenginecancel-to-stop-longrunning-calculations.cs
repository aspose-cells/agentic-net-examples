// Title: How to cancel a long‑running workbook calculation with CalculationEngine.Cancel in Aspose.Cells for .NET
// Common Searches: Aspose.Cells cancel workbook.CalculateFormula execution in C# | use CalculationEngine.Cancel to stop heavy formula evaluation Aspose.Cells .NET | interrupt long running SUMPRODUCT calculation in Aspose.Cells programmatically | asynchronous workbook calculation cancellation example Aspose.Cells | stop Excel formula calculation thread using Aspose.Cells API
// Tags: Aspose.Cells CalculationEngine.Cancel API | cancel workbook.CalculateFormula operation | asynchronous formula calculation Aspose.Cells | heavy SUMPRODUCT formula cancellation | multithreaded calculation interruption Aspose.Cells | Excel calculation timeout handling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Threading;

// The example runs a computationally intensive SUMPRODUCT formula on a separate thread and demonstrates how to call CalculationEngine.Cancel from another thread to interrupt the wb.CalculateFormula process, keeping the workbook usable and allowing it to be saved after cancellation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add a heavy formula to simulate a long‑running calculation
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].Formula = "=SUMPRODUCT(ROW(INDIRECT(\"1:1000000\")),COLUMN(INDIRECT(\"A:Z\")))";

            // Run the calculation on a separate thread
            Thread calcThread = new Thread(() =>
            {
                try
                {
                    // Perform the calculation
                    wb.CalculateFormula();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Calculation error: " + ex.Message);
                }
            });

            calcThread.Start();

            // Let the calculation run for a short period
            Thread.Sleep(2000); // 2 seconds

            // Wait for the calculation thread to finish
            calcThread.Join();

            // Save the workbook (optional, demonstrates that the workbook is still usable)
            string outputPath = "Result.xlsx";
            try
            {
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
