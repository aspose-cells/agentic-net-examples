// Title: Aspose.Cells C# – Custom Calculation Engine that Treats Empty Cells as Zero in SUM
// Description: This example shows how to subclass Aspose.Cells.AbstractCalculationEngine to intercept the SUM function, treat null or blank cells as 0, and return the correct total. All other functions are delegated to the default engine via the Skip flag. The custom engine is assigned to CalculationOptions.CustomEngine and used with Workbook.CalculateFormula to evaluate formulas that contain empty cells.
// Keywords: Aspose.Cells | C# | custom calculation engine | AbstractCalculationEngine | CalculationOptions.CustomEngine | treat empty cells as zero | SUM function override | blank cell handling | Excel formula customization | Aspose.Cells API
// Common Searches: Aspose.Cells treat blank cells as zero | custom calculation engine Aspose.Cells C# | override SUM function Aspose.Cells | CalculationOptions CustomEngine example | ignore empty cells in Aspose.Cells formulas | Aspose.Cells custom engine for SUM | C# Aspose.Cells custom formula behavior
// Developer Intent: Create a custom calculation engine that evaluates SUM formulas by counting empty cells as zero while leaving other functions unchanged.
// Use Cases: Generate financial reports where missing entries should be counted as zero. | Build dashboards that sum data ranges containing blanks without extra preprocessing. | Run server‑side .NET services that calculate workbooks with custom zero‑handling logic. | Apply the engine to legacy spreadsheets that use empty cells for optional values.
// AI Prompts: Provide C# code for an AbstractCalculationEngine subclass that treats null cell values as 0 in the SUM function and integrates it with CalculationOptions.CustomEngine. | Explain how to extend the custom engine to also handle the AVERAGE function while treating empty cells as zero. | Give troubleshooting steps when the Skip property is not accessible in a custom Aspose.Cells calculation engine.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomCalcOption
{
    // Custom calculation engine that treats empty cells as zero for the SUM function.
    // This example shows how to subclass Aspose.Cells.AbstractCalculationEngine to intercept the SUM function, treat null or blank cells as 0, and return the correct total. All other functions are delegated to the default engine via the Skip flag. The custom engine is assigned to CalculationOptions.CustomEngine and used with Workbook.CalculateFormula to evaluate formulas that contain empty cells.
    public class EmptyAsZeroEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions so we can intercept them.
        public override bool ProcessBuiltInFunctions => true;

        public override void Calculate(CalculationData data)
        {
            // Intercept only the SUM function; other functions should use the default engine.
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
                                // Treat null (empty cell) as zero.
                                sum += cellValue == null ? 0.0 : Convert.ToDouble(cellValue);
                            }
                        }
                    }
                    else
                    {
                        // Single scalar parameter.
                        sum += param == null ? 0.0 : Convert.ToDouble(param);
                    }
                }

                // Return the calculated sum.
                data.CalculatedValue = sum;
                return;
            }

            // For any other function, instruct the default engine to handle it.
            // The 'Skip' property may not be available in all versions, so set it via reflection.
            var skipProp = data.GetType().GetProperty("Skip");
            if (skipProp != null && skipProp.CanWrite)
            {
                skipProp.SetValue(data, true);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Populate sample data.
                ws.Cells["A1"].PutValue(10);   // Non‑empty cell.
                ws.Cells["B1"].PutValue(null); // Explicitly empty cell (treated as blank).

                // Formula that sums A1 and B1. B1 is empty.
                ws.Cells["C1"].Formula = "=SUM(A1:B1)";

                // Set up calculation options with the custom engine.
                CalculationOptions opts = new CalculationOptions
                {
                    CustomEngine = new EmptyAsZeroEngine(),
                    // Ensure errors are ignored so the demo runs smoothly.
                    IgnoreError = true,
                    Recursive = true
                };

                // Calculate all formulas using the custom options.
                wb.CalculateFormula(opts);

                // Output the result. Expected: 10 (since empty B1 is treated as 0).
                Console.WriteLine("Result of SUM(A1:B1) with empty cells as zero: " + ws.Cells["C1"].Value);

                // Save the workbook (optional, demonstrates lifecycle compliance).
                string outputPath = "CustomCalcOptionDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
