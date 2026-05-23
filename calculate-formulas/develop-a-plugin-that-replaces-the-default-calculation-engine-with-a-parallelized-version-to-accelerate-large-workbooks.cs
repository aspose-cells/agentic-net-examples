using System;
using System.Threading.Tasks;
using Aspose.Cells;

// Custom calculation engine that processes a heavy function in parallel
public class ParallelCalculationEngine : AbstractCalculationEngine
{
    // We only handle custom functions; built‑in functions are processed by the default engine
    public override bool ProcessBuiltInFunctions => false;

    public override void Calculate(CalculationData data)
    {
        // Custom function name: PARALLELSUM
        if (data.FunctionName.Equals("PARALLELSUM", StringComparison.OrdinalIgnoreCase))
        {
            // Expect a single range parameter
            if (data.ParamCount > 0)
            {
                object param = data.GetParamValue(0);
                double sum = 0;

                if (param is ReferredArea area)
                {
                    int rows = area.EndRow - area.StartRow + 1;
                    int cols = area.EndColumn - area.StartColumn + 1;
                    object lockObj = new object();

                    // Parallel loop over rows to sum numeric values
                    Parallel.For(0, rows, r =>
                    {
                        double localSum = 0;
                        for (int c = 0; c < cols; c++)
                        {
                            object val = area.GetValue(r, c);
                            if (val != null && double.TryParse(val.ToString(), out double d))
                            {
                                localSum += d;
                            }
                        }
                        lock (lockObj)
                        {
                            sum += localSum;
                        }
                    });
                }

                // Set the calculated result
                data.CalculatedValue = sum;
            }
        }
    }

    // No special force‑recalculation logic needed
    public override bool ForceRecalculate(string functionName) => false;
}

// Plugin entry point
public class ParallelEnginePlugin
{
    public static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("LargeData.xlsx"); // load

        // Example: add a formula that uses the parallel sum function
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["B1"].Formula = "=PARALLELSUM(A1:A10000)";

        // Configure calculation options to use the custom parallel engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = new ParallelCalculationEngine(),
            Recursive = true,
            IgnoreError = false
        };

        // Perform calculation with the parallel engine
        workbook.CalculateFormula(options);

        // Output the result to console
        Console.WriteLine("Result of PARALLELSUM: " + sheet.Cells["B1"].Value);

        // Save the workbook after calculation
        workbook.Save("LargeData_Calculated.xlsx"); // save
    }
}