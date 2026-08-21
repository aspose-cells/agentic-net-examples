// Title: Aspose.Cells C# – Fast MYFASTSUM with a Cached Custom Calculation Engine
// Description: Demonstrates how to speed up Cell.Calculate for large ranges by implementing a custom calculation engine that caches the result of the MYFASTSUM function, iterates cells using numeric values, disables recursive evaluation, and adjusts the calculation stack size. The workbook fills column A with 10,000 numbers, applies =MYFASTSUM(A1:A10000) in B1, runs the optimized engine, and saves the file.
// Keywords: Aspose.Cells custom engine | C# fast sum function | MYFASTSUM caching | optimize Cell.Calculate | large range summation .NET | non‑recursive calculation options | performance tuning Aspose.Cells
// Common Searches: Aspose.Cells cache custom function result | how to speed up MYFASTSUM in C# | disable recursive calculation Aspose.Cells | custom calculation engine example Aspose.Cells | optimize large range formulas .NET
// Developer Intent: Reduce the runtime of Cell.Calculate by creating a lightweight, cache‑aware custom engine for the MYFASTSUM function and by configuring calculation options for maximum throughput.
// Use Cases: Repeatedly sum the same massive range in a workbook without recomputing each call. | One‑time evaluation of complex formulas on worksheets with tens of thousands of rows. | Integrating a thread‑safe cache for any range‑based custom function in Aspose.Cells.
// AI Prompts: Write a thread‑safe version of the OptimizedEngine that supports parallel processing of range cells. | Generate benchmark code comparing the default Aspose.Cells engine with the cached OptimizedEngine for summing 20,000 cells. | Create a generic cached calculation engine template that works for SUM, AVERAGE, and custom functions in Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to speed up Cell.Calculate for large ranges by implementing a custom calculation engine that caches the result of the MYFASTSUM function, iterates cells using numeric values, disables recursive evaluation, and adjusts the calculation stack size. The workbook fills column A with 10,000 numbers, applies =MYFASTSUM(A1:A10000) in B1, runs the optimized engine, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate a large dataset (e.g., 10,000 rows) in column A
        for (int i = 0; i < 10000; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // A1, A2, ...
        }

        // Use a custom function MYFASTSUM that sums the whole range
        sheet.Cells["B1"].Formula = "=MYFASTSUM(A1:A10000)";

        // Configure calculation options with the optimized custom engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = new OptimizedEngine(),
            // Turn off recursive calculation when not needed
            Recursive = false,
            // Adjust stack size to avoid overflow while keeping performance
            CalcStackSize = 500
        };

        // Perform the calculation once
        workbook.CalculateFormula(options);

        // Output the result
        Console.WriteLine("Result of MYFASTSUM: " + sheet.Cells["B1"].Value);

        // Save the workbook (uses the standard Aspose.Cells save API)
        workbook.Save("OptimizedDemo.xlsx");
    }
}

// Custom calculation engine that caches results for identical parameter ranges
public class OptimizedEngine : AbstractCalculationEngine
{
    // Cache key: worksheet name + range address
    private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

    // Do not force recalculation for the custom function; allow caching
    public override bool ForceRecalculate(string functionName)
    {
        return false; // return true only for volatile functions
    }

    public override void Calculate(CalculationData data)
    {
        // Handle only the custom function MYFASTSUM; other functions fall back to default engine
        if (!string.Equals(data.FunctionName, "MYFASTSUM", StringComparison.OrdinalIgnoreCase))
            return;

        // Expect exactly one parameter (a range)
        if (data.ParamCount == 0)
        {
            data.CalculatedValue = "#VALUE!";
            return;
        }

        // Retrieve the parameter as a ReferredArea (range object)
        var area = data.GetParamValue(0) as ReferredArea;
        if (area == null)
        {
            data.CalculatedValue = "#VALUE!";
            return;
        }

        // Build a unique cache key based on worksheet and range coordinates
        string cacheKey = $"{data.Worksheet.Name}!{area.StartRow}:{area.StartColumn}-{area.EndRow}:{area.EndColumn}";

        // Return cached result if available
        if (_cache.TryGetValue(cacheKey, out object cachedResult))
        {
            data.CalculatedValue = cachedResult;
            return;
        }

        // Efficiently compute the sum by iterating cells directly
        double sum = 0;
        for (int r = area.StartRow; r <= area.EndRow; r++)
        {
            for (int c = area.StartColumn; c <= area.EndColumn; c++)
            {
                Cell cell = data.Worksheet.Cells[r, c];
                // Use the numeric value directly when possible to avoid extra conversions
                if (cell.Type == CellValueType.IsNumeric)
                {
                    sum += cell.DoubleValue;
                }
                else
                {
                    // Attempt to parse non‑numeric values that may represent numbers
                    if (double.TryParse(cell.StringValue, out double parsed))
                        sum += parsed;
                }
            }
        }

        // Cache the computed sum for future identical calls
        _cache[cacheKey] = sum;
        data.CalculatedValue = sum;
    }
}
