using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom calculation engine that implements the FISCALQTR function.
    // The function expects a single date argument and returns the fiscal quarter number.
    // Fiscal year is assumed to start in April:
    //   Q1 = Apr‑Jun, Q2 = Jul‑Sep, Q3 = Oct‑Dec, Q4 = Jan‑Mar.
    public class FiscalQuarterEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the called function is our custom one.
            if (data.FunctionName.Equals("FISCALQTR", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the first parameter value.
                object param = data.GetParamValue(0);

                DateTime date;

                // The parameter may be a ReferredArea (cell reference) or a direct value.
                if (param is ReferredArea area)
                {
                    // Get the underlying value from the referred cell.
                    object val = area.GetValue(0, 0);
                    date = ConvertToDate(val);
                }
                else
                {
                    date = ConvertToDate(param);
                }

                // Determine the fiscal quarter.
                int month = date.Month;
                int fiscalQuarter;

                // Fiscal year starts in April.
                if (month >= 4 && month <= 6)        // Apr, May, Jun
                    fiscalQuarter = 1;
                else if (month >= 7 && month <= 9)   // Jul, Aug, Sep
                    fiscalQuarter = 2;
                else if (month >= 10 && month <= 12) // Oct, Nov, Dec
                    fiscalQuarter = 3;
                else                                 // Jan, Feb, Mar
                    fiscalQuarter = 4;

                // Return the quarter as a double (Excel numeric type).
                data.CalculatedValue = (double)fiscalQuarter;
            }
        }

        // Helper to convert various possible Excel value types to DateTime.
        private DateTime ConvertToDate(object value)
        {
            if (value is DateTime dt)
                return dt;

            if (value is double oaDate)
                return DateTime.FromOADate(oaDate);

            if (value is string s && DateTime.TryParse(s, out DateTime parsed))
                return parsed;

            // Fallback to today if conversion fails.
            return DateTime.Today;
        }
    }

    public class FiscalQuarterDemo
    {
        public static void Run()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample dates.
            sheet.Cells["A1"].PutValue(new DateTime(2023, 4, 15)); // Expected Q1
            sheet.Cells["A2"].PutValue(new DateTime(2023, 7, 10)); // Expected Q2
            sheet.Cells["A3"].PutValue(new DateTime(2023, 11, 5)); // Expected Q3
            sheet.Cells["A4"].PutValue(new DateTime(2024, 2, 20)); // Expected Q4

            // Apply the custom fiscal quarter function to each date.
            sheet.Cells["B1"].Formula = "=FISCALQTR(A1)";
            sheet.Cells["B2"].Formula = "=FISCALQTR(A2)";
            sheet.Cells["B3"].Formula = "=FISCALQTR(A3)";
            sheet.Cells["B4"].Formula = "=FISCALQTR(A4)";

            // Set up calculation options with the custom engine.
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new FiscalQuarterEngine()
            };

            // Calculate all formulas in the workbook using the custom engine.
            workbook.CalculateFormula(options);

            // Output the results to the console.
            Console.WriteLine("Date\t\tFiscal Quarter");
            for (int i = 0; i < 4; i++)
            {
                string dateStr = sheet.Cells[i, 0].StringValue;
                string quarter = sheet.Cells[i, 1].StringValue;
                Console.WriteLine($"{dateStr}\t{quarter}");
            }

            // Save the workbook (uses the provided lifecycle rule for saving).
            workbook.Save("FiscalQuarterDemo.xlsx");
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            FiscalQuarterDemo.Run();
        }
    }
}