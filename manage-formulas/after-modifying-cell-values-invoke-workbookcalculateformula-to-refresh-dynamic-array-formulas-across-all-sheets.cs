// Title: C# – Refresh Dynamic Array Formulas After Updating Source Cells with Aspose.Cells
// Description: Demonstrates how to modify source ranges, apply a FILTER dynamic‑array formula, and use Workbook.RefreshDynamicArrayFormulas(true) together with Workbook.CalculateFormula() to recalculate spilled results across all worksheets before saving the workbook.
// Keywords: Aspose.Cells | RefreshDynamicArrayFormulas | dynamic array formula | FILTER function | C# .NET | Workbook.CalculateFormula | spilled array recalculation | Excel dynamic arrays | Aspose.Cells example | recalculate formulas after data change
// Common Searches: Aspose.Cells refresh dynamic array after data change | C# refresh spilled array results Aspose.Cells | Workbook.RefreshDynamicArrayFormulas usage | CalculateFormula vs RefreshDynamicArrayFormulas | dynamic array FILTER example Aspose.Cells .NET | update source cells and recalc formulas Aspose.Cells
// Developer Intent: Ensure that dynamic‑array formulas reflect updated source data by refreshing them and then recalculating any other formulas in the workbook.
// Use Cases: After bulk editing source cells, call RefreshDynamicArrayFormulas(true) to update FILTER or SORT spills before exporting the file. | Combine RefreshDynamicArrayFormulas with CalculateFormula to keep both spilled array results and regular formulas synchronized in multi‑sheet reports. | Automate data‑import pipelines where source tables change frequently, guaranteeing accurate dynamic‑array outputs for downstream calculations.
// AI Prompts: Generate C# code that updates multiple source ranges, then invokes RefreshDynamicArrayFormulas and CalculateFormula to refresh all dynamic‑array and regular formulas in a workbook. | Explain when to use Workbook.RefreshDynamicArrayFormulas(true) versus Workbook.CalculateFormula in Aspose.Cells, with code snippets. | Provide a sample that iterates through every worksheet in a workbook and refreshes dynamic‑array formulas after a bulk data import using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayRefreshDemo
{
    // Demonstrates how to modify source ranges, apply a FILTER dynamic‑array formula, and use Workbook.RefreshDynamicArrayFormulas(true) together with Workbook.CalculateFormula() to recalculate spilled results across all worksheets before saving the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Populate source data that will be used by a dynamic array formula
            // ------------------------------------------------------------
            // Values in column A (A2:A5)
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(20);
            cells["A4"].PutValue(30);
            cells["A5"].PutValue(40);

            // Values in column B (B2:B5)
            cells["B2"].PutValue(5);
            cells["B3"].PutValue(15);
            cells["B4"].PutValue(25);
            cells["B5"].PutValue(35);

            // ------------------------------------------------------------
            // 2. Set a dynamic array formula that depends on the above data
            //    The formula will spill into column C (C2:C5)
            // ------------------------------------------------------------
            // Example: FILTER values from column A where column B > 20
            string dynamicFormula = "=FILTER(A2:A5, B2:B5>20)";
            // SetDynamicArrayFormula(string, FormulaParseOptions, bool) rule
            cells["C2"].SetDynamicArrayFormula(dynamicFormula, new FormulaParseOptions(), true);

            // ------------------------------------------------------------
            // 3. Change a source value that influences the dynamic array result
            // ------------------------------------------------------------
            // Update B3 from 15 to 30, which should now include A3 in the spill range
            cells["B3"].PutValue(30);

            // ------------------------------------------------------------
            // 4. Refresh dynamic array formulas after the data change
            //    RefreshDynamicArrayFormulas(bool) rule
            // ------------------------------------------------------------
            // 'true' indicates that the formulas should be recalculated as well
            workbook.RefreshDynamicArrayFormulas(true);

            // ------------------------------------------------------------
            // 5. Calculate all other formulas in the workbook (if any)
            //    CalculateFormula() rule
            // ------------------------------------------------------------
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 6. Output the spilled results to the console for verification
            // ------------------------------------------------------------
            Console.WriteLine("Spilled results of the dynamic array formula (C2:C5):");
            for (int row = 2; row <= 5; row++)
            {
                Cell cell = cells[row - 1, 2]; // Column C index = 2
                Console.WriteLine($"C{row}: {(cell.IsFormula ? cell.Value?.ToString() ?? "Empty" : "Not a formula")}");
            }

            // ------------------------------------------------------------
            // 7. Save the workbook (lifecycle save rule)
            // ------------------------------------------------------------
            workbook.Save("DynamicArrayRefreshResult.xlsx");
        }
    }
}
