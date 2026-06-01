using System;
using Aspose.Cells;

namespace AsposeCellsCustomCalcOption
{
    // Custom calculation engine that treats empty cells as zero for built‑in functions.
    public class EmptyAsZeroEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions so that this engine can override them.
        public override bool ProcessBuiltInFunctions => true;

        public override void Calculate(CalculationData data)
        {
            // Handle the SUM function as an example.
            // Other functions can be added similarly.
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0.0;

                // Iterate over all parameters passed to the function.
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // Parameter can be a single value or a ReferredArea (range).
                    if (param is ReferredArea area)
                    {
                        // Loop through each cell in the range.
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object cellValue = area.GetValue(r, c);

                                // Treat null, empty string, or DBNull as zero.
                                if (cellValue == null || cellValue == string.Empty || cellValue == DBNull.Value)
                                {
                                    sum += 0.0;
                                }
                                else
                                {
                                    // Convert numeric values; non‑numeric values are ignored.
                                    if (double.TryParse(cellValue.ToString(), out double d))
                                        sum += d;
                                }
                            }
                        }
                    }
                    else
                    {
                        // Single scalar parameter.
                        if (param == null || param == string.Empty || param == DBNull.Value)
                        {
                            sum += 0.0;
                        }
                        else if (double.TryParse(param.ToString(), out double d))
                        {
                            sum += d;
                        }
                    }
                }

                // Set the calculated result for the SUM function.
                data.CalculatedValue = sum;
                return;
            }

            // For any other function, let the default engine handle it.
            // Since ProcessBuiltInFunctions is true, we must explicitly skip calculation
            // for functions we do not handle to avoid incorrect results.
            SkipCalculation();
        }
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: A1 = 5, A2 = empty, A3 = 10
            sheet.Cells["A1"].PutValue(5);
            // A2 left empty intentionally
            sheet.Cells["A3"].PutValue(10);

            // Set a formula that sums the range, which includes an empty cell.
            sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

            // ---------- Define custom calculation options ----------
            CalculationOptions options = new CalculationOptions
            {
                // Attach the custom engine that treats empty cells as zero.
                CustomEngine = new EmptyAsZeroEngine()
            };

            // ---------- Perform calculation ----------
            workbook.CalculateFormula(options);

            // Output the result of the formula.
            Console.WriteLine("Result of SUM(A1:A3) with empty cells treated as zero: " + sheet.Cells["B1"].Value);

            // ---------- Save the workbook ----------
            workbook.Save("CustomCalcOptionDemo.xlsx");
        }
    }
}