using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomEngineTests
{
    // Custom calculation engine that overrides the built‑in SUM function.
    public class DoubleSumEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions.
        public override bool ProcessBuiltInFunctions => true;

        public override void Calculate(CalculationData data)
        {
            // Handle only the SUM function; other functions fall back to the default engine.
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;
                // Iterate through all parameters of the SUM function.
                for (int i = 0; i < data.ParamCount; i++)
                {
                    // Parameters are supplied as ReferredArea objects for range arguments.
                    if (data.GetParamValue(i) is ReferredArea area)
                    {
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object cellValue = area.GetValue(r, c);
                                if (cellValue != null && double.TryParse(cellValue.ToString(), out double d))
                                    sum += d;
                            }
                        }
                    }
                    else
                    {
                        // Single scalar arguments (e.g., numbers) are also possible.
                        object arg = data.GetParamValue(i);
                        if (arg != null && double.TryParse(arg.ToString(), out double d))
                            sum += d;
                    }
                }

                // Custom logic: double the calculated sum.
                data.CalculatedValue = sum * 2;
            }
        }

        public override bool ForceRecalculate(string functionName) => false;
    }

    // Custom calculation engine that does NOT process built‑in functions.
    public class NoBuiltInEngine : AbstractCalculationEngine
    {
        public override bool ProcessBuiltInFunctions => false;

        public override void Calculate(CalculationData data)
        {
            // No custom handling; let the default engine compute everything.
        }

        public override bool ForceRecalculate(string functionName) => false;
    }

    internal static class SimpleAssert
    {
        public static void AreEqual(double expected, double actual, string message = "")
        {
            const double tolerance = 1e-9;
            if (Math.Abs(expected - actual) > tolerance)
                throw new Exception($"Assert Failed: Expected {expected}, Actual {actual}. {message}");
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                TestSumIsOverriddenWhenProcessBuiltInFunctionsIsTrue();
                TestSumIsNotOverriddenWhenProcessBuiltInFunctionsIsFalse();
                TestWorkbookCanBeSavedAfterCustomCalculation();

                Console.WriteLine("All tests passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        static void TestSumIsOverriddenWhenProcessBuiltInFunctionsIsTrue()
        {
            // Arrange
            var wb = new Workbook();
            var ws = wb.Worksheets[0];

            ws.Cells["A1"].PutValue(1);
            ws.Cells["A2"].PutValue(2);
            ws.Cells["A3"].PutValue(3);

            // SUM should be overridden (6 * 2 = 12)
            ws.Cells["B1"].Formula = "=SUM(A1:A3)";
            // AVERAGE should remain unchanged ( (1+2+3)/3 = 2 )
            ws.Cells["B2"].Formula = "=AVERAGE(A1:A3)";

            var options = new CalculationOptions { CustomEngine = new DoubleSumEngine() };

            // Act
            wb.CalculateFormula(options);

            // Assert
            SimpleAssert.AreEqual(12.0, Convert.ToDouble(ws.Cells["B1"].Value), "SUM override failed.");
            SimpleAssert.AreEqual(2.0, Convert.ToDouble(ws.Cells["B2"].Value), "AVERAGE should be unchanged.");
        }

        static void TestSumIsNotOverriddenWhenProcessBuiltInFunctionsIsFalse()
        {
            // Arrange
            var wb = new Workbook();
            var ws = wb.Worksheets[0];

            ws.Cells["A1"].PutValue(4);
            ws.Cells["A2"].PutValue(5);
            ws.Cells["A3"].PutValue(6);

            ws.Cells["B1"].Formula = "=SUM(A1:A3)";      // Expected 15
            ws.Cells["B2"].Formula = "=AVERAGE(A1:A3)"; // Expected 5

            var options = new CalculationOptions { CustomEngine = new NoBuiltInEngine() };

            // Act
            wb.CalculateFormula(options);

            // Assert
            SimpleAssert.AreEqual(15.0, Convert.ToDouble(ws.Cells["B1"].Value), "SUM should not be overridden.");
            SimpleAssert.AreEqual(5.0, Convert.ToDouble(ws.Cells["B2"].Value), "AVERAGE calculation failed.");
        }

        static void TestWorkbookCanBeSavedAfterCustomCalculation()
        {
            // Arrange
            var wb = new Workbook();
            var ws = wb.Worksheets[0];

            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].Formula = "=SUM(A1:A2)";

            var options = new CalculationOptions { CustomEngine = new DoubleSumEngine() };
            wb.CalculateFormula(options);

            // Act & Assert (no exception should be thrown)
            string outputPath = "CustomEngineResult.xlsx";

            try
            {
                // Save the workbook; wrap in try‑catch for safety.
                wb.Save(outputPath);
                // Verify file was created.
                if (!File.Exists(outputPath))
                    throw new Exception("File was not saved as expected.");
            }
            catch (Exception ex)
            {
                throw new Exception($"Saving workbook failed: {ex.Message}");
            }
            finally
            {
                // Clean up the generated file to avoid side effects.
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }
    }
}