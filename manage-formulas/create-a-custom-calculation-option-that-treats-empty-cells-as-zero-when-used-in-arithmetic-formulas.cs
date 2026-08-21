// Title: Aspose.Cells C# Custom Calculation Engine – Treat Blank Cells as Zero in Formulas
// Description: Shows how to build a custom Aspose.Cells calculation engine that treats empty or blank cells as 0 when evaluating arithmetic expressions and the SUM function, and how to apply it via CalculationOptions.CustomEngine.
// Keywords: Aspose.Cells custom engine | treat blank cells as zero | C# calculation engine Aspose | override SUM function | CalculationOptions CustomEngine | handle empty cells in formulas | .NET spreadsheet calculation | Aspose.Cells AbstractCalculationEngine
// Common Searches: Aspose.Cells treat empty cells as zero | custom calculation engine C# Aspose.Cells | override SUM function Aspose.Cells | CalculationOptions.CustomEngine example | blank cell handling in Aspose.Cells formulas
// Developer Intent: Create a custom calculation engine that substitutes null or empty cell values with zero during formula evaluation.
// Use Cases: Accurately sum ranges that contain optional or missing data without manual cleanup. | Generate financial or statistical reports where blank entries must be counted as zero. | Run bulk spreadsheet calculations in .NET while preventing conversion errors from empty cells.
// AI Prompts: Write a C# class extending AbstractCalculationEngine that treats empty cells as zero for all arithmetic functions, not just SUM. | Show how to configure CalculationOptions to use a custom engine and calculate every formula in an Aspose.Cells workbook. | Explain how to extend the TreatEmptyAsZeroEngine to also handle the AVERAGE function by ignoring blanks or treating them as zero.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomCalcOption
{
    // Custom calculation engine that treats empty cells as zero for the SUM function.
    // It also processes built‑in functions so that the engine is invoked for SUM.
    // Shows how to build a custom Aspose.Cells calculation engine that treats empty or blank cells as 0 when evaluating arithmetic expressions and the SUM function, and how to apply it via CalculationOptions.CustomEngine.
    public class TreatEmptyAsZeroEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions.
        public override bool ProcessBuiltInFunctions => true;

        public override void Calculate(CalculationData data)
        {
            // Only customize the SUM function; other functions fall back to default behavior.
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0.0;

                // Iterate over all parameters passed to SUM.
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // Parameter can be a single value or a ReferredArea (range).
                    if (param is ReferredArea area)
                    {
                        // Walk through each cell in the range.
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object cellValue = area.GetValue(r, c);
                                // Treat null or empty as zero.
                                sum += ConvertToDoubleOrZero(cellValue);
                            }
                        }
                    }
                    else
                    {
                        // Single scalar parameter.
                        sum += ConvertToDoubleOrZero(param);
                    }
                }

                // Set the calculated result for the SUM function.
                data.CalculatedValue = sum;
            }
            // For any other function we do nothing – the default engine will handle it.
        }

        // Helper: converts a value to double; null or empty becomes 0.
        private static double ConvertToDoubleOrZero(object value)
        {
            if (value == null) return 0.0;
            if (value is string s && string.IsNullOrWhiteSpace(s)) return 0.0;
            try
            {
                return Convert.ToDouble(value);
            }
            catch
            {
                // If conversion fails, treat as zero.
                return 0.0;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate data: A1 = 10, A2 = empty, A3 = 5
            ws.Cells["A1"].PutValue(10);
            // A2 left empty intentionally
            ws.Cells["A3"].PutValue(5);

            // Formula that adds the three cells. Empty cell should be treated as zero.
            ws.Cells["B1"].Formula = "=A1+A2+A3";

            // Also demonstrate SUM over a range that includes an empty cell.
            ws.Cells["B2"].Formula = "=SUM(A1:A3)";

            // ---------- Set up custom calculation options ----------
            CalculationOptions opts = new CalculationOptions
            {
                // Attach the custom engine.
                CustomEngine = new TreatEmptyAsZeroEngine(),
                // Ensure recursive calculation (default true) and ignore errors.
                Recursive = true,
                IgnoreError = true
            };

            // ---------- Perform calculation ----------
            // Calculate all formulas in the workbook using the custom options.
            wb.CalculateFormula(opts);

            // ---------- Output results ----------
            Console.WriteLine("Result of A1+A2+A3 (B1): " + ws.Cells["B1"].Value); // Expected 15
            Console.WriteLine("Result of SUM(A1:A3) (B2): " + ws.Cells["B2"].Value); // Expected 15

            // ---------- Save the workbook (optional) ----------
            wb.Save("CustomCalcOptionDemo.xlsx");
        }
    }
}
