// Title: Custom FISCALQUARTER UDF in Aspose.Cells for .NET – Return Fiscal Quarter from a Date
// Description: Demonstrates how to build a user‑defined function called FISCALQUARTER by extending Aspose.Cells' AbstractCalculationEngine. The engine extracts a single date argument (cell reference, Excel serial number, or string), converts it to DateTime, computes the calendar quarter, and returns the quarter as a numeric value. The example creates a workbook, inserts a sample date, applies =FISCALQUARTER(A1), configures CalculationOptions with the custom engine, calculates formulas, prints the result, and saves the file.
// Keywords: Aspose.Cells | C# | custom function | user defined function | UDF | FiscalQuarter | quarter calculation | AbstractCalculationEngine | Excel formula | date to quarter | financial reporting | custom calculation engine | Excel serial date | OADate
// Common Searches: Aspose.Cells custom function example | How to create user defined function in Aspose.Cells C# | Fiscal quarter function Aspose.Cells | Calculate quarter from date using Aspose.Cells | Custom calculation engine Aspose.Cells .NET
// Developer Intent: Implement a user‑defined Excel function named FISCALQUARTER that returns the fiscal quarter for a given date using Aspose.Cells for .NET.
// Use Cases: Generate quarterly financial reports by converting transaction dates to fiscal quarters without altering source data. | Create dynamic dashboards that group sales or expenses by quarter using the =FISCALQUARTER formula across multiple worksheets. | Automate KPI aggregation where quarter‑level summaries are required for budgeting or forecasting. | Support custom fiscal calendars by extending the function to shift the quarter start month.
// AI Prompts: Write a C# class that extends AbstractCalculationEngine to add a user‑defined function FISCALQUARTER handling cell references, OLE Automation dates, and string dates, returning the calendar quarter as a double. | Provide sample code that sets CalculationOptions.CustomEngine to the custom engine, inserts a date into cell A1, uses =FISCALQUARTER(A1) in cell B1, calculates the workbook, prints the results, and saves the file. | Explain how to modify the FISCALQUARTER function to support a fiscal year that starts in July instead of January. | Generate unit tests for the FiscalQuarterEngine covering valid dates, invalid inputs, and edge cases such as leap years.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom calculation engine that implements a user‑defined function FISCALQUARTER(date)
    // Demonstrates how to build a user‑defined function called FISCALQUARTER by extending Aspose.Cells' AbstractCalculationEngine. The engine extracts a single date argument (cell reference, Excel serial number, or string), converts it to DateTime, computes the calendar quarter, and returns the quarter as a numeric value. The example creates a workbook, inserts a sample date, applies =FISCALQUARTER(A1), configures CalculationOptions with the custom engine, calculates formulas, prints the result, and saves the file.
    public class FiscalQuarterEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the called function is our custom one (case‑insensitive)
            if (string.Equals(data.FunctionName, "FISCALQUARTER", StringComparison.OrdinalIgnoreCase))
            {
                // Expect exactly one argument – the date value
                if (data.ParamCount == 1)
                {
                    object param = data.GetParamValue(0);
                    DateTime date;

                    // The argument may be a ReferredArea (cell reference) or a direct double (Excel serial date)
                    if (param is ReferredArea area)
                    {
                        // Get the underlying value from the referenced cell
                        object val = area.GetValue(0, 0);
                        date = ConvertToDate(val);
                    }
                    else
                    {
                        // Direct value (likely a double representing OLE Automation date)
                        date = ConvertToDate(param);
                    }

                    // Compute fiscal quarter (standard calendar quarters)
                    int quarter = ((date.Month - 1) / 3) + 1;

                    // Return the quarter as a numeric value
                    data.CalculatedValue = (double)quarter;
                    return;
                }

                // If the argument count is wrong, return a NaN to indicate an error
                data.CalculatedValue = double.NaN;
                return;
            }

            // For any other function, defer to the base implementation (or leave unhandled)
        }

        // Helper to convert various possible Excel value types to DateTime
        private DateTime ConvertToDate(object value)
        {
            if (value is DateTime dt)
                return dt;

            // Excel stores dates as double (OLE Automation date)
            if (value is double d)
                return DateTime.FromOADate(d);

            // Try parsing a string representation
            if (value is string s && DateTime.TryParse(s, out DateTime parsed))
                return parsed;

            // Fallback to today if conversion fails
            return DateTime.Today;
        }
    }

    public class FiscalQuarterFunctionDemo
    {
        public static void Run()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Place a sample date in cell A1 (e.g., 2023‑08‑15)
                sheet.Cells["A1"].PutValue(new DateTime(2023, 8, 15));

                // Use the custom function in cell B1
                sheet.Cells["B1"].Formula = "=FISCALQUARTER(A1)";

                // Set up calculation options to use our custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new FiscalQuarterEngine()
                };

                // Calculate all formulas in the workbook using the custom engine
                workbook.CalculateFormula(options);

                // Output the result to the console
                Console.WriteLine("Date in A1: " + sheet.Cells["A1"].StringValue);
                Console.WriteLine("Fiscal Quarter (B1): " + sheet.Cells["B1"].StringValue);

                // ---------- Save the workbook ----------
                workbook.Save("FiscalQuarterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            FiscalQuarterFunctionDemo.Run();
        }
    }
}
