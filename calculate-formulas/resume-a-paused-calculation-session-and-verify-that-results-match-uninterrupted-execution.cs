using System;
using Aspose.Cells;

class ResumeCalculationDemo
{
    static void Main()
    {
        // Create a workbook with data and formulas
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        for (int i = 0; i < 1000; i++)
        {
            ws.Cells[i, 0].PutValue(i);               // Column A
            ws.Cells[i, 1].PutValue(i * 2);           // Column B
            ws.Cells[i, 2].Formula = $"=A{i + 1}+B{i + 1}"; // Column C = A+B
        }

        // Save the original workbook (unmodified)
        wb.Save("Original.xlsx");

        // Load the workbook with an interrupt monitor that will pause calculation
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);
        LoadOptions loadOptions = new LoadOptions { InterruptMonitor = monitor };
        Workbook wbInterrupted = new Workbook("Original.xlsx", loadOptions);

        // Start monitoring with a very short time limit to force interruption
        monitor.StartMonitor(10); // 10 milliseconds

        try
        {
            // Attempt to calculate; this should be interrupted
            wbInterrupted.CalculateFormula();
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Calculation was interrupted as expected.");
        }

        // Remove the interrupt monitor and resume calculation
        wbInterrupted.InterruptMonitor = null;
        wbInterrupted.CalculateFormula(); // Resume full calculation

        // Save the workbook after resumed calculation
        wbInterrupted.Save("Resumed.xlsx");

        // Perform uninterrupted calculation on a fresh copy for verification
        Workbook wbFull = new Workbook("Original.xlsx");
        wbFull.CalculateFormula();
        wbFull.Save("FullCalculated.xlsx");

        // Compare a subset of cells to verify that results match
        bool allMatch = true;
        for (int i = 0; i < 10; i++)
        {
            object resumedValue = wbInterrupted.Worksheets[0].Cells[i, 2].Value;
            object fullValue = wbFull.Worksheets[0].Cells[i, 2].Value;
            if (!object.Equals(resumedValue, fullValue))
            {
                allMatch = false;
                Console.WriteLine($"Mismatch at row {i + 1}: resumed={resumedValue}, full={fullValue}");
            }
        }

        Console.WriteLine(allMatch
            ? "Resumed results match uninterrupted calculation."
            : "Results differ between resumed and uninterrupted calculations.");
    }
}