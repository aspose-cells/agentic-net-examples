using System;
using Aspose.Cells;

namespace CustomCalculationEngineDemo
{
    // Custom engine that provides implementations for user‑defined functions.
    // If a function is not recognized, it substitutes it with a built‑in equivalent
    // or returns a custom result.
    public class SubstituteEngine : AbstractCalculationEngine
    {
        // Example: indicate that we do not need parameters in array mode or literal text.
        public override bool IsParamArrayModeRequired => false;
        public override bool IsParamLiteralRequired => false;
        public override bool ProcessBuiltInFunctions => false;

        // Force recalculation for volatile custom functions (optional).
        public override bool ForceRecalculate(string functionName)
        {
            // Recalculate every time for functions that depend on external state.
            return functionName.Equals("NOWCUSTOM", StringComparison.OrdinalIgnoreCase);
        }

        public override void Calculate(CalculationData data)
        {
            string func = data.FunctionName?.ToUpperInvariant();

            if (func == "MYSUM")
            {
                // MYSUM(a,b,...) => sum of all numeric parameters
                double sum = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object val = data.GetParamValue(i);
                    if (val is double d) sum += d;
                    else if (double.TryParse(Convert.ToString(val), out d)) sum += d;
                }
                data.CalculatedValue = sum;
            }
            else if (func == "MYAVG")
            {
                // MYAVG(a,b,...) => average of numeric parameters
                double sum = 0;
                int count = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object val = data.GetParamValue(i);
                    if (val is double d) { sum += d; count++; }
                    else if (double.TryParse(Convert.ToString(val), out d)) { sum += d; count++; }
                }
                data.CalculatedValue = count > 0 ? sum / count : "#DIV/0!";
            }
            else if (func == "NOWCUSTOM")
            {
                // Substitute NOWCUSTOM with current date‑time
                data.CalculatedValue = DateTime.Now;
            }
            else
            {
                // Function not recognized – return #NAME? to let Excel handle it or provide fallback.
                data.CalculatedValue = "#NAME?";
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate some sample data.
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].PutValue(30);

            // Use custom functions in formulas.
            ws.Cells["B1"].Formula = "=MYSUM(A1, A2, A3)";   // Expected 60
            ws.Cells["B2"].Formula = "=MYAVG(A1, A2, A3)";   // Expected 20
            ws.Cells["B3"].Formula = "=NOWCUSTOM()";        // Expected current date‑time
            ws.Cells["B4"].Formula = "=UNKNOWNFUNC(1,2)";   // Will return #NAME?

            // Set calculation options to use our custom engine.
            CalculationOptions opts = new CalculationOptions
            {
                CustomEngine = new SubstituteEngine(),
                Recursive = true,
                IgnoreError = false
            };

            // Perform calculation.
            wb.CalculateFormula(opts);

            // Output results to console.
            Console.WriteLine("B1 (MYSUM)   = " + ws.Cells["B1"].Value);
            Console.WriteLine("B2 (MYAVG)   = " + ws.Cells["B2"].Value);
            Console.WriteLine("B3 (NOWCUSTOM)= " + ws.Cells["B3"].Value);
            Console.WriteLine("B4 (UNKNOWN) = " + ws.Cells["B4"].StringValue);

            // Save the workbook.
            wb.Save("SubstituteEngineResult.xlsx");
        }
    }
}