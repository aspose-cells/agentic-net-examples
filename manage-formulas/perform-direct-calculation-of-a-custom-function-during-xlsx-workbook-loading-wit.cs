using System;
using Aspose.Cells;

namespace DirectCustomFunctionCalculation
{
    // Custom calculation engine that implements a user‑defined function MYFUNC
    class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function MYFUNC (case‑insensitive)
            if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Expect exactly two numeric parameters
                if (data.ParamCount == 2)
                {
                    // Retrieve the evaluated parameter values
                    object p0 = data.GetParamValue(0);
                    object p1 = data.GetParamValue(1);

                    // Convert to double (Aspose.Cells may return double, int, etc.)
                    double v0 = Convert.ToDouble(p0);
                    double v1 = Convert.ToDouble(p1);

                    // Example logic: return the product of the two numbers
                    data.CalculatedValue = v0 * v1;
                }
                else
                {
                    // Incorrect number of parameters – return an error value
                    data.CalculatedValue = "#VALUE!";
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a temporary workbook and save it (simulating an existing file)
            // -----------------------------------------------------------------
            Workbook tempWb = new Workbook();
            tempWb.Worksheets[0].Name = "DataSheet";
            tempWb.Save("sample.xlsx", SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 2. Load the workbook from disk (the point where we want to calculate)
            // -----------------------------------------------------------------
            Workbook wb = new Workbook("sample.xlsx");

            // Obtain the first worksheet – it will be used only as a calculation context
            Worksheet sheet = wb.Worksheets[0];

            // -----------------------------------------------------------------
            // 3. Prepare calculation options with the custom engine
            // -----------------------------------------------------------------
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // -----------------------------------------------------------------
            // 4. Directly calculate a custom formula without inserting it into any cell
            // -----------------------------------------------------------------
            string formula = "=MYFUNC(7, 3)"; // Expected result: 21
            object result = sheet.CalculateFormula(formula, options);

            // -----------------------------------------------------------------
            // 5. Output the result
            // -----------------------------------------------------------------
            Console.WriteLine($"Result of {formula} = {result}");

            // -----------------------------------------------------------------
            // 6. (Optional) Save the workbook if further processing is needed
            // -----------------------------------------------------------------
            wb.Save("sample_processed.xlsx", SaveFormat.Xlsx);
        }
    }
}