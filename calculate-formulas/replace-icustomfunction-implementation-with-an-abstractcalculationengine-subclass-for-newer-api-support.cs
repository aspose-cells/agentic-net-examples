// Title: Replace ICustomFunction with an AbstractCalculationEngine subclass to implement MYFUNC in Aspose.Cells for .NET
// Common Searches: Aspose.Cells replace ICustomFunction with AbstractCalculationEngine example C# | create custom user-defined function MYFUNC using CalculationOptions.CustomEngine in .NET | C# calculate workbook formulas with a custom AbstractCalculationEngine subclass | how to sum scalar arguments and range values in a custom Aspose.Cells function | using ReferredArea inside a custom calculation engine in Aspose.Cells
// Tags: AbstractCalculationEngine subclass for custom formulas | Aspose.Cells custom calculation engine C# | CalculationOptions.CustomEngine usage | user-defined function MYFUNC implementation | handling ReferredArea in Aspose.Cells custom engine

using System;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // The example shows how to replace the legacy ICustomFunction interface by subclassing AbstractCalculationEngine. A MyCustomEngine class overrides Calculate to evaluate the MYFUNC function, summing numeric arguments and the first cell of any ReferredArea. The engine is assigned to CalculationOptions.CustomEngine, a workbook is populated, the formula =MYFUNC(A1,A2,A3) is set, and workbook.CalculateFormula computes and prints the result.
    class Program
    {
        // Helper to safely convert an object to double (treat non‑numeric as 0)
        private static double ConvertToDouble(object obj)
        {
            if (obj == null)
                return 0.0;

            if (obj is double d)
                return d;

            if (obj is int i)
                return i;

            if (double.TryParse(obj.ToString(), out double result))
                return result;

            return 0.0;
        }

        // Custom calculation engine that implements MYFUNC
        private class MyCustomEngine : AbstractCalculationEngine
        {
            public override void Calculate(CalculationData data)
            {
                try
                {
                    if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
                    {
                        double sum = 0.0;

                        for (int i = 0; i < data.ParamCount; i++)
                        {
                            object param = data.GetParamValue(i);

                            if (param is ReferredArea area)
                            {
                                // Take the first cell's value from the referred area
                                object val = area.GetValue(0, 0);
                                sum += ConvertToDouble(val);
                            }
                            else
                            {
                                sum += ConvertToDouble(param);
                            }
                        }

                        // Set the calculated result
                        data.CalculatedValue = sum;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error in custom engine: " + ex.Message);
                }
            }
        }

        static void Main(string[] args)
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Prepare a workbook with sample data and a custom formula
                // ------------------------------------------------------------
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate cells used as arguments
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Insert a formula that uses the custom function MYFUNC
                sheet.Cells["B1"].Formula = "=MYFUNC(A1,A2,A3)";

                // ------------------------------------------------------------
                // 2. Set up calculation options with the custom engine
                // ------------------------------------------------------------
                var options = new CalculationOptions
                {
                    CustomEngine = new MyCustomEngine()
                };

                // ------------------------------------------------------------
                // 3. Perform calculation using the custom engine
                // ------------------------------------------------------------
                workbook.CalculateFormula(options);

                // ------------------------------------------------------------
                // 4. Output the result
                // ------------------------------------------------------------
                Console.WriteLine("Result of MYFUNC(A1,A2,A3): " + sheet.Cells["B1"].Value);
            }
            catch (Exception ex)
            {
                // Any unexpected errors are reported here
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
