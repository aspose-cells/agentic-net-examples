using System;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsPowerQueryDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook that contains Power Query formulas (default XLSX format)
            Workbook workbook = new Workbook("source.xlsx");

            // Access the mashup data
            DataMashup mashup = workbook.DataMashup;

            // Ensure mashup and formulas collection are available
            if (mashup != null && mashup.PowerQueryFormulas != null && mashup.PowerQueryFormulas.Count > 0)
            {
                Console.WriteLine($"Found {mashup.PowerQueryFormulas.Count} Power Query formula(s).");

                // Iterate through each Power Query formula
                foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                {
                    // Display current formula details
                    Console.WriteLine($"Formula Name: {formula.Name}");
                    Console.WriteLine($"Definition: {formula.FormulaDefinition}");

                    // Example modification: prepend a comment indicating default XLSX usage
                    string modifiedDefinition = "// Extracted using default XLSX format\n" + formula.FormulaDefinition;
                    
                    // PowerQueryFormula.FormulaDefinition is read‑only, but we can modify its items if needed.
                    // Here we update the first item value if items exist.
                    if (formula.PowerQueryFormulaItems != null && formula.PowerQueryFormulaItems.Count > 0)
                    {
                        PowerQueryFormulaItem firstItem = formula.PowerQueryFormulaItems[0];
                        firstItem.Value = modifiedDefinition;
                        Console.WriteLine("Modified first item value to include default XLSX comment.");
                    }
                }
            }
            else
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
            }

            // Save the workbook (still in default XLSX format)
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
    }
}