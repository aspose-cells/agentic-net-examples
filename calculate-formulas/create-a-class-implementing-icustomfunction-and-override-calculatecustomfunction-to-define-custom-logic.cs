using System;
using System.IO;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // Custom calculation engine that implements the "MYFUNC" function.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            try
            {
                // Handle only the custom function named "MYFUNC".
                if (string.Equals(data.FunctionName, "MYFUNC", StringComparison.OrdinalIgnoreCase) &&
                    data.ParamCount >= 2)
                {
                    // Retrieve and convert the first two parameters to double.
                    double val0 = Convert.ToDouble(data.GetParamValue(0));
                    double val1 = Convert.ToDouble(data.GetParamValue(1));

                    // Return the sum.
                    data.CalculatedValue = val0 + val1;
                }
                else
                {
                    // Not enough parameters or different function – return #VALUE! error.
                    data.CalculatedValue = "#VALUE!";
                }
            }
            catch
            {
                // Conversion failed – return #VALUE! error.
                data.CalculatedValue = "#VALUE!";
            }
        }

        // Force recalculation for the custom function (volatile behavior).
        public override bool ForceRecalculate(string functionName)
        {
            return string.Equals(functionName, "MYFUNC", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and obtain the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data.
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(25);

                // Set a formula that uses the custom function MYFUNC.
                sheet.Cells["B1"].Formula = "=MYFUNC(A1, A2)";

                // Configure calculation options to use the custom engine.
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyCustomEngine()
                };

                // Perform the calculation.
                workbook.CalculateFormula(options);

                // Output the result.
                Console.WriteLine("Result of MYFUNC(A1, A2): " + sheet.Cells["B1"].Value);

                // Define output file path.
                string outputPath = "CustomFunctionResult.xlsx";

                // Ensure the directory exists before saving.
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}