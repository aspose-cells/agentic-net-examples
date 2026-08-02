// Title: Custom WORKDAYS function in Aspose.Cells for .NET – count business days in formulas
// Description: Demonstrates how to create a WorkDaysEngine that inherits AbstractCalculationEngine, intercepts the WORKDAYS function, converts Excel serial dates to DateTime, swaps out‑of‑order dates, iterates over the range counting only Monday‑Friday days, and returns the total as a double. The engine is attached to CalculationOptions, the formula "=WORKDAYS(A1,B1)" is placed in a cell, wb.CalculateFormula evaluates it, and the result is printed and saved to an XLSX file.
// Keywords: Aspose.Cells custom function | WORKDAYS .NET | business days calculation | custom calculation engine C# | Excel formula extension | date serial conversion | weekday count algorithm
// Common Searches: Aspose.Cells how to add a custom WORKDAYS function | calculate weekdays between two dates in C# using Aspose.Cells | implement custom Excel function with AbstractCalculationEngine | count business days in a workbook formula | extend Aspose.Cells with user‑defined functions
// Developer Intent: Create a reusable WORKDAYS user‑defined function that returns the number of Monday‑to‑Friday days between two dates for use in Aspose.Cells formulas.
// Use Cases: Enable end‑users to enter start and end dates in cells and obtain the business‑day count with =WORKDAYS(start,end). | Integrate the custom engine into existing calculation pipelines so all formulas, including the new function, are evaluated automatically. | Persist the calculated result by saving the workbook after calling wb.CalculateFormula with the custom engine.
// AI Prompts: Write C# code that registers a custom calculation engine in Aspose.Cells to implement a WORKDAYS function that excludes weekends. | Show how to call the custom WORKDAYS function from a worksheet formula and retrieve its numeric result programmatically. | Explain how to extend WorkDaysEngine to accept an optional holiday range and subtract those dates from the business‑day total.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements the WORKDAYS function
    // Demonstrates how to create a WorkDaysEngine that inherits AbstractCalculationEngine, intercepts the WORKDAYS function, converts Excel serial dates to DateTime, swaps out‑of‑order dates, iterates over the range counting only Monday‑Friday days, and returns the total as a double. The engine is attached to CalculationOptions, the formula "=WORKDAYS(A1,B1)" is placed in a cell, wb.CalculateFormula evaluates it, and the result is printed and saved to an XLSX file.
    public class WorkDaysEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is our custom function
            if (data.FunctionName != null && data.FunctionName.Equals("WORKDAYS", StringComparison.OrdinalIgnoreCase))
            {
                // Ensure we have at least two parameters (start date and end date)
                if (data.ParamCount < 2)
                {
                    data.CalculatedValue = 0;
                    return;
                }

                // Retrieve the first two parameters (Excel serial numbers)
                object startObj = data.GetParamValue(0);
                object endObj   = data.GetParamValue(1);

                // Convert parameters to DateTime. Excel stores dates as double (OADate)
                DateTime startDate = Convert.ToDouble(startObj) == 0 ? DateTime.MinValue : DateTime.FromOADate(Convert.ToDouble(startObj));
                DateTime endDate   = Convert.ToDouble(endObj)   == 0 ? DateTime.MinValue : DateTime.FromOADate(Convert.ToDouble(endObj));

                // If start date is after end date, swap them
                if (startDate > endDate)
                {
                    var temp = startDate;
                    startDate = endDate;
                    endDate = temp;
                }

                int workDays = 0;
                for (DateTime dt = startDate; dt <= endDate; dt = dt.AddDays(1))
                {
                    // Weekday values: Monday = 1, ..., Sunday = 0 (or 7)
                    DayOfWeek dow = dt.DayOfWeek;
                    if (dow != DayOfWeek.Saturday && dow != DayOfWeek.Sunday)
                    {
                        workDays++;
                    }
                }

                // Return the count as a double (Aspose.Cells expects numeric results)
                data.CalculatedValue = (double)workDays;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate start and end dates (Excel serial dates are automatically handled)
            cells["A1"].PutValue(new DateTime(2023, 7, 24)); // Monday
            cells["B1"].PutValue(new DateTime(2023, 7, 31)); // Next Monday

            // Set the custom formula using the WORKDAYS function
            cells["C1"].Formula = "=WORKDAYS(A1,B1)";

            // Prepare calculation options with our custom engine
            CalculationOptions opts = new CalculationOptions
            {
                CustomEngine = new WorkDaysEngine()
            };

            // Calculate all formulas in the workbook using the custom engine
            wb.CalculateFormula(opts);

            // Output the result of the custom function
            Console.WriteLine("Working days between A1 and B1: " + cells["C1"].Value);

            // Save the workbook (lifecycle rule: save)
            wb.Save("WorkDaysCustomFunctionDemo.xlsx");
        }
    }
}
