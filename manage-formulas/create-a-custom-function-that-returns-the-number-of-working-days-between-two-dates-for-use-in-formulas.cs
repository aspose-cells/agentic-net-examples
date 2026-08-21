// Title: C# – Add a custom WORKDAYS function in Aspose.Cells to count business days
// Description: Shows how to create a user‑defined WORKDAYS function in Aspose.Cells for .NET by extending AbstractCalculationEngine. The sample inserts start and end dates, applies =WORKDAYS(A1,B1) in a cell, converts parameters, swaps dates if needed, excludes Saturdays and Sundays, returns the weekday count, calculates the workbook, and saves the file.
// Keywords: Aspose.Cells | custom function | WORKDAYS | business days | C# | .NET | AbstractCalculationEngine | user defined function | Excel formula | calculate weekdays | NETWORKDAYS alternative | date calculation | GitHub example
// Common Searches: Aspose.Cells custom WORKDAYS function | C# calculate business days with Aspose.Cells | how to add user defined function in Aspose.Cells | Aspose.Cells count weekdays between dates | implement NETWORKDAYS equivalent in .NET
// Developer Intent: Create a user‑defined WORKDAYS function that returns the number of weekdays between two dates inside an Aspose.Cells workbook.
// Use Cases: Project scheduling – compute duration while ignoring weekends | Payroll processing – determine work days in a pay period | Resource planning – calculate billable days for staff | Replace Excel's NETWORKDAYS in server‑side .NET reports | Automate SLA calculations that exclude weekend days
// AI Prompts: Write C# code that registers a WORKDAYS UDF in Aspose.Cells using AbstractCalculationEngine, handling date conversion and weekend exclusion. | Show how to call =WORKDAYS(A1,B1) in a worksheet and retrieve the result after wb.CalculateFormula. | Extend the WORKDAYS engine to accept an optional holiday list as a third parameter and skip those dates. | Generate unit tests for the custom WORKDAYS function covering reversed dates, same‑day inputs, and edge cases.

using System;
using Aspose.Cells;

// Shows how to create a user‑defined WORKDAYS function in Aspose.Cells for .NET by extending AbstractCalculationEngine. The sample inserts start and end dates, applies =WORKDAYS(A1,B1) in a cell, converts parameters, swaps dates if needed, excludes Saturdays and Sundays, returns the weekday count, calculates the workbook, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Sample start and end dates (Excel stores dates as OLE Automation dates)
        cells["A1"].PutValue(new DateTime(2023, 7, 24)); // Monday
        cells["B1"].PutValue(new DateTime(2023, 7, 31)); // Next Monday

        // Use the custom function WORKDAYS in a formula
        cells["C1"].Formula = "=WORKDAYS(A1,B1)";

        // Set calculation options to use the custom engine
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = new WorkDaysEngine()
        };

        // Calculate all formulas in the workbook using the custom engine
        wb.CalculateFormula(opts);

        // Output the result of the custom function
        Console.WriteLine("Working days between A1 and B1: " + cells["C1"].Value);

        // Save the workbook (lifecycle rule: use provided save method)
        wb.Save("WorkDaysDemo.xlsx");
    }

    // Custom calculation engine that implements the WORKDAYS function
    class WorkDaysEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is our custom function
            if (string.Equals(data.FunctionName, "WORKDAYS", StringComparison.OrdinalIgnoreCase))
            {
                // Ensure we have at least two parameters (start date, end date)
                if (data.ParamCount >= 2)
                {
                    object startObj = data.GetParamValue(0);
                    object endObj   = data.GetParamValue(1);

                    DateTime startDate = ConvertToDate(startObj);
                    DateTime endDate   = ConvertToDate(endObj);

                    // If dates are reversed, swap them
                    if (startDate > endDate)
                    {
                        DateTime tmp = startDate;
                        startDate = endDate;
                        endDate = tmp;
                    }

                    // Compute working days (exclude Saturday and Sunday)
                    int workDays = CountWorkDays(startDate, endDate);
                    data.CalculatedValue = workDays;
                }
                else
                {
                    // Not enough parameters – return 0
                    data.CalculatedValue = 0;
                }
            }
        }

        // Helper to convert Excel parameter values to DateTime
        private DateTime ConvertToDate(object value)
        {
            if (value is double d)               // Excel serial date
                return DateTime.FromOADate(d);
            if (value is DateTime dt)            // Already a DateTime
                return dt;
            if (value != null && DateTime.TryParse(value.ToString(), out DateTime parsed))
                return parsed;                    // Parse from string if possible
            return DateTime.MinValue;            // Fallback
        }

        // Helper to count weekdays between two dates inclusive
        private int CountWorkDays(DateTime start, DateTime end)
        {
            int count = 0;
            for (DateTime d = start; d <= end; d = d.AddDays(1))
            {
                if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday)
                    count++;
            }
            return count;
        }
    }
}
