// Title: Resume a paused workbook calculation and verify identical results with Aspose.Cells for .NET
// AI Prompts: Show how to use SystemTimeInterruptMonitor to pause a workbook's formula evaluation, then continue the evaluation and retrieve the resulting values in C#. | Provide C# code that checks whether the values obtained after continuing a calculation match those from a prior uninterrupted run, using Aspose.Cells.
// Common Searches: Aspose.Cells how to interrupt and resume formula calculation in C# | C# continue workbook calculation after SystemTimeInterruptMonitor interruption | verify that resumed formula results equal original calculation Aspose.Cells | example of using interrupt monitor to pause Excel calculation with Aspose.Cells | compare interrupted calculation results with full calculation Aspose.Cells .NET
// Tags: Aspose.Cells SystemTimeInterruptMonitor usage | resume workbook calculation Aspose.Cells | C# verify formula results after continue | compare interrupted vs full calculation Aspose.Cells | calculate formulas with interrupt monitor .NET

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationResumeDemo
{
    // The sample creates a workbook, fills column A with numbers and column B with formulas, runs a full calculation and stores the results, then deliberately interrupts a second calculation using SystemTimeInterruptMonitor. After catching the interruption, the monitor is restarted with a longer timeout, the calculation is continued, and each resumed cell value is compared to the original full‑run value to ensure they match before saving the workbook.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and populate data + formulas
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];

            // Fill column A with numbers 1..1000
            for (int i = 0; i < 1000; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);
                // Column B formula: =A1*2 (will be copied down)
                sheet.Cells[i, 1].Formula = $"=A{i + 1}*2";
            }

            // -------------------------------------------------
            // 2. Perform uninterrupted calculation and store results
            // -------------------------------------------------
            workbook.CalculateFormula(); // full calculation

            double[] fullResults = new double[1000];
            for (int i = 0; i < 1000; i++)
            {
                fullResults[i] = sheet.Cells[i, 1].DoubleValue;
            }

            // -------------------------------------------------
            // 3. Set up an interrupt monitor to pause calculation
            // -------------------------------------------------
            // Use SystemTimeInterruptMonitor with a very short limit to force interruption
            SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(terminateWithoutException: false);
            workbook.InterruptMonitor = monitor; // assign to workbook

            // Start monitor with 5 ms limit (likely to interrupt)
            monitor.StartMonitor(5);

            try
            {
                // Re‑calculate; this should be interrupted
                workbook.CalculateFormula();
                Console.WriteLine("Unexpected: calculation completed without interruption.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
            {
                Console.WriteLine("Calculation was successfully interrupted.");
            }

            // -------------------------------------------------
            // 4. Resume calculation after interruption
            // -------------------------------------------------
            // Give a generous time limit so the remaining calculation can finish
            monitor.StartMonitor(10000); // 10 seconds

            // Resume calculation
            workbook.CalculateFormula();

            // -------------------------------------------------
            // 5. Verify that resumed results match the original full calculation
            // -------------------------------------------------
            bool allMatch = true;
            for (int i = 0; i < 1000; i++)
            {
                double resumedValue = sheet.Cells[i, 1].DoubleValue;
                if (Math.Abs(resumedValue - fullResults[i]) > 1e-9)
                {
                    allMatch = false;
                    Console.WriteLine($"Mismatch at row {i + 1}: full={fullResults[i]}, resumed={resumedValue}");
                    break;
                }
            }

            Console.WriteLine(allMatch
                ? "Verification succeeded: resumed results match uninterrupted calculation."
                : "Verification failed: results differ.");

            // -------------------------------------------------
            // 6. Save the workbook (using the same lifecycle rule)
            // -------------------------------------------------
            workbook.Save("ResumedCalculationResult.xlsx"); // save
        }
    }
}
