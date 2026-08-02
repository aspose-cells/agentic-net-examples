// Title: C# ClearCache method for Aspose.Cells custom calculation engine to reset cached formula results
// Description: Shows how to extend AbstractCalculationEngine with a public ClearCache() method that empties the in‑memory dictionary used by a custom SUMCACHED function, guaranteeing fresh computation after workbook data changes.
// Keywords: Aspose.Cells | C# | .NET | custom calculation engine | cache invalidation | ClearCache method | reset cached formulas | SUMCACHED function | ForceRecalculate | workbook recalculation | in‑memory cache
// Common Searches: Aspose.Cells clear custom engine cache | reset cached custom function C# | how to invalidate calculation cache Aspose.Cells | ClearCache example Aspose.Cells | force recalc custom function Aspose.Cells
// Developer Intent: Add a public method that clears the internal cache of a custom Aspose.Cells calculation engine so formulas are recomputed with up‑to‑date cell values.
// Use Cases: After modifying source cells, call engine.ClearCache() before wb.CalculateFormula(opts) to obtain correct SUMCACHED results. | In a long‑running service that processes many workbooks, invoke ClearCache periodically to free memory and avoid stale values. | Hook engine.ClearCache() to a workbook change event to automatically invalidate cached custom function results.
// AI Prompts: Generate C# code that implements a ClearCache() method for an Aspose.Cells custom calculation engine and demonstrates its usage in a workbook recalculation workflow. | Explain the interaction between ForceRecalculate and ClearCache in a custom engine and outline best practices for cache invalidation. | Create a sample that clears the custom engine cache, updates cell values, recalculates, and verifies the new formula output.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Shows how to extend AbstractCalculationEngine with a public ClearCache() method that empties the in‑memory dictionary used by a custom SUMCACHED function, guaranteeing fresh computation after workbook data changes.
public class CachedCustomEngine : AbstractCalculationEngine
{
    // Simple in‑memory cache for demonstration purposes
    private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

    public override void Calculate(CalculationData data)
    {
        // Build a cache key based on function name and number of parameters
        string key = data.FunctionName + "_" + data.ParamCount;

        // Return cached result if it exists
        if (_cache.TryGetValue(key, out var cachedValue))
        {
            data.CalculatedValue = cachedValue;
            return;
        }

        // Example custom function: SUMCACHED
        if (data.FunctionName.Equals("SUMCACHED", StringComparison.OrdinalIgnoreCase))
        {
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                // Parameters are passed as ReferredArea objects
                ReferredArea area = (ReferredArea)data.GetParamValue(i);
                sum += Convert.ToDouble(area.GetValue(0, 0));
            }

            data.CalculatedValue = sum;
            _cache[key] = sum; // Store result in cache
        }
        else
        {
            // For other functions let the default engine handle them
            data.CalculatedValue = data.GetParamValue(0);
        }
    }

    // Force recalculation for the custom function so that changes in source data are detected
    public override bool ForceRecalculate(string functionName)
    {
        return functionName.Equals("SUMCACHED", StringComparison.OrdinalIgnoreCase);
    }

    // Public method to clear the internal cache, ensuring fresh computation after data changes
    public void ClearCache()
    {
        _cache.Clear();
    }
}

public class CustomEngineCacheDemo
{
    public static void Test()
    {
        try
        {
            // Create a new workbook and populate some data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue(5);
            ws.Cells["A2"].PutValue(10);
            ws.Cells["A3"].Formula = "=SUMCACHED(A1,A2)";

            // Instantiate the custom engine with caching
            CachedCustomEngine engine = new CachedCustomEngine();

            // Set calculation options to use the custom engine
            CalculationOptions opts = new CalculationOptions { CustomEngine = engine };

            // First calculation – cache will be populated
            wb.CalculateFormula(opts);
            Console.WriteLine("First result: " + ws.Cells["A3"].Value); // Expected 15

            // Modify source data
            ws.Cells["A1"].PutValue(20);

            // Recalculate – ForceRecalculate forces a new calculation, but cache would still be used if not cleared
            wb.CalculateFormula(opts);
            Console.WriteLine("After data change without clearing cache: " + ws.Cells["A3"].Value); // Expected 30

            // Clear the engine's cache explicitly
            engine.ClearCache();

            // Recalculate again – now the engine recomputes without any cached value
            wb.CalculateFormula(opts);
            Console.WriteLine("After clearing cache: " + ws.Cells["A3"].Value); // Expected 30

            // Ensure output directory exists before saving
            string outputPath = "CustomEngineCacheDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during the demo: " + ex.Message);
        }
    }
}

public class Program
{
    public static void Main()
    {
        CustomEngineCacheDemo.Test();
    }
}
