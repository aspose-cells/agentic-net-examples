using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace CustomFunctionInvalidationDemo
{
    // Custom calculation engine that forces recalculation of the custom function
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Force recalculation for the custom function "MYVOLATILEFUNC"
        public override bool ForceRecalculate(string functionName)
        {
            return string.Equals(functionName, "MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase);
        }

        // Simple implementation: return the value of the first parameter plus 1
        public override void Calculate(CalculationData data)
        {
            if (string.Equals(data.FunctionName, "MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Get the first parameter (could be a cell reference, a constant, etc.)
                object param = data.GetParamValue(0);
                double value = 0;

                // If the parameter is a ReferredArea (cell/range), extract its first value
                if (param is ReferredArea area && area.StartRow == area.EndRow && area.StartColumn == area.EndColumn)
                {
                    value = Convert.ToDouble(area.GetValue(0, 0));
                }
                else if (param != null)
                {
                    value = Convert.ToDouble(param);
                }

                // Example calculation: add 1 to the input value
                data.CalculatedValue = value + 1;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create ----------
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Put a source value that will be used by the custom function
                cells["B1"].PutValue(10);

                // Set a formula that uses the custom function with a cell reference
                cells["A1"].Formula = "=MYVOLATILEFUNC(B1)";

                // Prepare calculation options with the custom engine
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CustomEngine = new MyCustomEngine()
                };

                // Initial calculation – the custom function is evaluated
                workbook.CalculateFormula(calcOptions);
                Console.WriteLine($"Initial A1 value (should be 11): {cells["A1"].Value}");

                // ---------- Modify dependent cell ----------
                // Change the value that the custom function depends on
                cells["B1"].PutValue(20);

                // Retrieve all cells that depend on B1 (including A1)
                IEnumerator dependents = cells.GetDependentsInCalculation(0, 1, true); // row 0, column 1 => B1
                Console.WriteLine("Cells that depend on B1:");
                while (dependents.MoveNext())
                {
                    Cell depCell = (Cell)dependents.Current;
                    Console.WriteLine($"- {depCell.Name}");
                }

                // Recalculate the workbook; ForceRecalculate ensures the cached result is ignored
                workbook.CalculateFormula(calcOptions);
                Console.WriteLine($"After B1 change, A1 value (should be 21): {cells["A1"].Value}");

                // ---------- Save ----------
                string outputPath = "CustomFunctionInvalidationDemo.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}