// Title: Configure a Calculation Timeout in Aspose.Cells for .NET with a Custom AbstractCalculationMonitor
// Description: This example shows how to create a TimeoutMonitor that inherits from AbstractCalculationMonitor, sets a deadline based on a millisecond limit, and throws an exception in BeforeCalculate when the limit is exceeded. The monitor is attached to CalculationOptions.CalculationMonitor, used in workbook.CalculateFormula inside a try‑catch block, and the workbook is then saved. The approach prevents long‑running formulas from hanging applications such as web services.
// Keywords: Aspose.Cells | C# | Calculation timeout | AbstractCalculationMonitor | CalculationOptions | Formula evaluation limit | prevent hanging formulas | Excel processing performance | exception handling | web API Excel
// Common Searches: Aspose.Cells set formula calculation timeout | C# custom calculation monitor Aspose.Cells | How to abort long formulas in Aspose.Cells | Timeout for workbook.CalculateFormula | Prevent Excel calculation hang using Aspose.Cells
// Developer Intent: Add a configurable time limit to formula calculation so that any formula exceeding the limit aborts with an exception.
// Use Cases: Enforce a 2‑second limit when processing workbooks that may contain heavy formulas or user‑defined functions. | Keep an ASP.NET web service responsive by terminating long‑running calculations on uploaded Excel files. | Apply a global calculation timeout across multiple workbooks in a batch‑processing job.
// AI Prompts: Generate C# code for a custom AbstractCalculationMonitor that stops calculation after a specified number of milliseconds. | Show how to attach a timeout monitor to CalculationOptions and handle the timeout exception gracefully. | Explain how to log the sheet, row, and column where the timeout occurs inside BeforeCalculate.

using System;
using Aspose.Cells;

// This example shows how to create a TimeoutMonitor that inherits from AbstractCalculationMonitor, sets a deadline based on a millisecond limit, and throws an exception in BeforeCalculate when the limit is exceeded. The monitor is attached to CalculationOptions.CalculationMonitor, used in workbook.CalculateFormula inside a try‑catch block, and the workbook is then saved. The approach prevents long‑running formulas from hanging applications such as web services.
class Program
{
    // Custom calculation monitor that aborts when the specified time limit is exceeded
    class TimeoutMonitor : AbstractCalculationMonitor
    {
        private readonly DateTime _deadline;

        public TimeoutMonitor(int timeoutMilliseconds)
        {
            // Calculate the moment when the timeout will occur
            _deadline = DateTime.UtcNow.AddMilliseconds(timeoutMilliseconds);
        }

        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // If the current time passed the deadline, abort the calculation
            if (DateTime.UtcNow > _deadline)
                throw new Exception("Formula calculation timed out.");
        }

        // Called after each cell is calculated (no action needed)
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex) { }

        // Handle circular references (continue calculation)
        public override bool OnCircular(System.Collections.IEnumerator circularCellsData) => true;
    }

    static void Main()
    {
        // Create a new workbook and add some sample data / formulas
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["B1"].Formula = "=SUM(A1:A2)";          // Simple formula
        sheet.Cells["B2"].Formula = "=POWER(2,30)";        // Potentially heavy calculation

        // Set up calculation options with the custom timeout monitor (e.g., 2 seconds)
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = new TimeoutMonitor(2000) // 2000 ms timeout
        };

        try
        {
            // Perform calculation; will be interrupted if it exceeds the timeout
            workbook.CalculateFormula(calcOptions);
            Console.WriteLine("Calculation completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Calculation interrupted: " + ex.Message);
        }

        // Save the workbook (the save operation itself is not timed out here)
        workbook.Save("Result.xlsx");
    }
}
