// Title: Implement a formula caching engine in C# with Aspose.Cells to skip duplicate calculations
// Common Searches: c# Aspose.Cells cache formula evaluation to improve performance | how to avoid repeated formula calculations in an Aspose.Cells workbook | custom formula cache implementation using Aspose.Cells CalculationOptions | evaluate all cells in an Excel file with caching using Aspose.Cells C# | store and reuse formula results in Aspose.Cells to reduce calculation time
// Tags: Aspose.Cells formula result caching | C# custom calculation engine for Excel workbooks | duplicate formula evaluation avoidance | Excel workbook calculation optimization with Aspose.Cells | Aspose.Cells CalculationOptions caching strategy

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCacheEngine
{
    // Engine that caches results of identical formula strings
    // The example shows how to build a FormulaCacheEngine that records the results of identical formula strings in a dictionary and reuses them during evaluation. It iterates through each worksheet and cell, calculates formulas only when they haven't been cached, writes the evaluated value back, and saves the workbook. This approach reduces redundant calculations and speeds up processing of large Excel files using Aspose.Cells in C#.
    class FormulaCacheEngine
    {
        private readonly Workbook _workbook;
        // Cache: formula text -> evaluated result
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>();

        public FormulaCacheEngine(Workbook workbook)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        // Evaluate a single cell, using cache when possible
        public object EvaluateCell(Cell cell)
        {
            try
            {
                // Non‑formula cells just return their current value
                if (!cell.IsFormula)
                    return cell.Value;

                string formulaKey = cell.Formula;

                // Return cached result if the same formula was already evaluated
                if (_cache.TryGetValue(formulaKey, out object cachedResult))
                    return cachedResult;

                // Perform actual calculation using Aspose.Cells built‑in evaluator
                var options = new CalculationOptions(); // default options
                cell.Calculate(options);               // Calculate writes the result into the cell
                object result = cell.Value;             // Retrieve the calculated value

                // Store result in cache for future identical formulas
                _cache[formulaKey] = result;
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error evaluating cell {cell.Name}: {ex.Message}");
                return cell.Value; // Fallback to original value
            }
        }

        // Iterate through all worksheets and cells, evaluating formulas with caching
        public void EvaluateAll()
        {
            foreach (Worksheet sheet in _workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        object value = EvaluateCell(cell);
                        // Write the evaluated value back to the cell (optional)
                        cell.PutValue(value);
                    }
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists; otherwise create an empty workbook
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                    workbook = new Workbook(); // creates a default workbook
                }

                // Instantiate the caching engine
                var engine = new FormulaCacheEngine(workbook);

                // Evaluate all formulas, avoiding redundant calculations
                engine.EvaluateAll();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
