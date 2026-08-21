// Title: Create a custom FISCALQUARTER function in Aspose.Cells for .NET
// Description: Demonstrates how to extend Aspose.Cells.AbstractCalculationEngine to implement a user‑defined FISCALQUARTER function. The engine validates a single date argument (DateTime, Excel OLE Automation double, or ReferredArea), shifts the month according to a fiscal year start month (April), computes the fiscal quarter, and returns it as a numeric value. The example shows registering the engine via CalculationOptions, applying the formula =FISCALQUARTER(A1) in a worksheet, calculating formulas, and saving the workbook.
// Keywords: Aspose.Cells custom function | Fiscal quarter C# | AbstractCalculationEngine example | Excel user‑defined function .NET | Fiscal year start month | Date to quarter conversion | C# Excel formula extension
// Common Searches: custom fiscal quarter function Aspose.Cells | how to create user defined function in Aspose.Cells C# | calculate fiscal quarter from date .NET | Aspose.Cells AbstractCalculationEngine tutorial | handle Excel serial dates in custom Aspose.Cells function
// Developer Intent: Add a user‑defined function that returns the fiscal quarter for a given date, respecting a configurable fiscal year start month, and use it in workbook formulas.
// Use Cases: Insert =FISCALQUARTER(date) in financial reports to derive quarter numbers automatically. | Apply the custom engine across multiple worksheets or named ranges that contain transaction dates. | Generate quarterly summaries by dragging the FISCALQUARTER formula down a column of dates.
// AI Prompts: Write a C# class that inherits Aspose.Cells.AbstractCalculationEngine and implements a FISCALQUARTER function returning the fiscal quarter based on a date, with the fiscal year starting in April. | Show how to register the custom calculation engine in CalculationOptions and use the =FISCALQUARTER formula in a worksheet. | Explain how to detect and convert DateTime, Excel OLE Automation serial numbers, and ReferredArea parameters inside a custom Aspose.Cells function.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements the FISCALQUARTER function
    // Demonstrates how to extend Aspose.Cells.AbstractCalculationEngine to implement a user‑defined FISCALQUARTER function. The engine validates a single date argument (DateTime, Excel OLE Automation double, or ReferredArea), shifts the month according to a fiscal year start month (April), computes the fiscal quarter, and returns it as a numeric value. The example shows registering the engine via CalculationOptions, applying the formula =FISCALQUARTER(A1) in a worksheet, calculating formulas, and saving the workbook.
    public class FiscalQuarterEngine : AbstractCalculationEngine
    {
        // Define the fiscal year start month (e.g., April = 4)
        private const int FiscalYearStartMonth = 4;

        public override void Calculate(CalculationData data)
        {
            // Check if the function being called is our custom function
            if (data.FunctionName.Equals("FISCALQUARTER", StringComparison.OrdinalIgnoreCase))
            {
                // Expect exactly one parameter (a date)
                if (data.ParamCount != 1)
                {
                    // Return a #VALUE! error (using NaN as a placeholder)
                    data.CalculatedValue = double.NaN;
                    return;
                }

                // Retrieve the parameter value
                object param = data.GetParamValue(0);
                DateTime date;

                // Parameter may be a DateTime, a double (Excel serial date), or a ReferredArea
                if (param is DateTime dt)
                {
                    date = dt;
                }
                else if (param is double oaDate)
                {
                    // Excel stores dates as OLE Automation dates
                    date = DateTime.FromOADate(oaDate);
                }
                else if (param is ReferredArea area)
                {
                    // Get the first cell value from the referred area
                    object val = area.GetValue(0, 0);
                    if (val is DateTime dt2)
                    {
                        date = dt2;
                    }
                    else if (val is double oaDate2)
                    {
                        date = DateTime.FromOADate(oaDate2);
                    }
                    else
                    {
                        data.CalculatedValue = double.NaN;
                        return;
                    }
                }
                else
                {
                    data.CalculatedValue = double.NaN;
                    return;
                }

                // Compute fiscal quarter
                int month = date.Month;
                // Shift month based on fiscal year start
                int shiftedMonth = (month - FiscalYearStartMonth + 12) % 12 + 1;
                int quarter = ((shiftedMonth - 1) / 3) + 1;

                // Return quarter as a numeric value
                data.CalculatedValue = (double)quarter;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put a sample date in A1 (e.g., 2023-05-15)
                cells["A1"].PutValue(new DateTime(2023, 5, 15));

                // Use the custom function in B1
                cells["B1"].Formula = "=FISCALQUARTER(A1)";

                // Set up calculation options with the custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new FiscalQuarterEngine()
                };

                // Calculate all formulas using the custom engine
                sheet.CalculateFormula(options, true);

                // Output the result to console
                Console.WriteLine($"Date in A1: {cells["A1"].StringValue}");
                Console.WriteLine($"Fiscal Quarter (custom function) in B1: {cells["B1"].StringValue}");

                // Save the workbook
                workbook.Save("FiscalQuarterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
