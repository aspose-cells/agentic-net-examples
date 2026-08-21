// Title: Resume Interrupted Formula Calculation with Aspose.Cells .NET and Verify Results
// Description: Demonstrates how to pause a long‑running calculation using ThreadInterruptMonitor, resume it, and compare the resumed result with an uninterrupted reference workbook in C#.
// Keywords: Aspose.Cells resume calculation | ThreadInterruptMonitor .NET | interrupt and continue formula evaluation | verify calculation result | C# Aspose.Cells example
// Common Searches: how to resume a paused calculation in Aspose.Cells | Aspose.Cells ThreadInterruptMonitor usage | compare interrupted vs uninterrupted formula results | C# resume workbook calculation after timeout | Aspose.Cells verify resumed calculation
// Developer Intent: Pause a formula calculation, resume it later, and ensure the final value matches the original uninterrupted computation.
// Use Cases: Free CPU resources by interrupting a heavy calculation and completing it when resources are available. | Validate data integrity after a calculation restart by comparing with a pre‑computed reference. | Automate workbook saving only after successful resumption and verification of formulas.
// AI Prompts: Write C# code that uses Aspose.Cells ThreadInterruptMonitor to interrupt a calculation after 5 ms and then resume it with a longer timeout. | Create a method that catches the CellsException for an interrupted calculation, restarts the monitor, and checks the result against a reference workbook. | Explain best practices for handling interrupted calculations in Aspose.Cells, covering monitor configuration, exception handling, and result verification.

using System;
using Aspose.Cells;

namespace AsposeCellsResumeCalculationDemo
{
    // Demonstrates how to pause a long‑running calculation using ThreadInterruptMonitor, resume it, and compare the resumed result with an uninterrupted reference workbook in C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a workbook with sample data and formulas
                // -------------------------------------------------
                Workbook referenceWb = new Workbook();
                Worksheet refSheet = referenceWb.Worksheets[0];

                // Fill column A with numbers 1..1000
                for (int i = 0; i < 1000; i++)
                    refSheet.Cells[i, 0].PutValue(i + 1);

                // B1 = SUM(A1:A1000)
                refSheet.Cells["B1"].Formula = "=SUM(A1:A1000)";

                // Perform uninterrupted calculation
                referenceWb.CalculateFormula();

                // Store expected result
                double expectedResult = refSheet.Cells["B1"].DoubleValue;
                Console.WriteLine($"Expected result (uninterrupted): {expectedResult}");

                // -------------------------------------------------
                // 2. Create a second workbook with identical data
                // -------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];

                for (int i = 0; i < 1000; i++)
                    sheet.Cells[i, 0].PutValue(i + 1);

                sheet.Cells["B1"].Formula = "=SUM(A1:A1000)";

                // -------------------------------------------------
                // 3. Set up an interrupt monitor to pause calculation
                // -------------------------------------------------
                ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(terminateWithoutException: false);
                wb.InterruptMonitor = monitor;

                // Start monitor with a very short time limit to force interruption
                monitor.StartMonitor(5); // 5 ms

                try
                {
                    // This calculation will be interrupted
                    wb.CalculateFormula();
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
                {
                    Console.WriteLine("Calculation was interrupted as expected.");
                }

                // -------------------------------------------------
                // 4. Resume calculation by starting a longer monitor
                // -------------------------------------------------
                monitor.StartMonitor(10000); // 10 seconds – enough to finish

                // Resume calculation
                wb.CalculateFormula();

                // -------------------------------------------------
                // 5. Verify that the resumed result matches the reference
                // -------------------------------------------------
                double resumedResult = sheet.Cells["B1"].DoubleValue;
                Console.WriteLine($"Resumed result: {resumedResult}");

                if (Math.Abs(resumedResult - expectedResult) < 1e-9)
                    Console.WriteLine("Verification succeeded: results match.");
                else
                    Console.WriteLine("Verification failed: results differ.");

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                string outputPath = "ResumedCalculationResult.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // General exception handling to avoid unexpected crashes
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
