// Title: C# – Apply UNIQUE Dynamic Array Formula with Aspose.Cells to Auto‑Deduplicate Column C into Column D
// Description: Demonstrates how to create a workbook, fill C1:C6 with duplicate fruit names, set the dynamic array formula "=UNIQUE(C1:C6)" in D1 using SetDynamicArrayFormula, calculate and refresh the spill range, read the unique values from column D, and save the file as UniqueDynamicArrayDemo.xlsx.
// Keywords: Aspose.Cells UNIQUE function | dynamic array formula C# | SetDynamicArrayFormula | deduplicate column values | spill range refresh | CalculateFormula Aspose.Cells | RefreshDynamicArrayFormulas | .NET spreadsheet automation
// Common Searches: how to use UNIQUE with Aspose.Cells .NET | set dynamic array formula in C# Aspose.Cells | auto‑deduplicate column using UNIQUE function | refresh spilled array results Aspose.Cells | read UNIQUE spill range programmatically
// Developer Intent: Insert a UNIQUE dynamic‑array formula in D1 so that distinct values from C1:C6 automatically spill into column D, then retrieve and save the results.
// Use Cases: Create a live list of distinct product names for dashboards. | Generate a summary of unique customer IDs that updates with source data changes. | Export a clean set of categories to another sheet without manual filtering.
// AI Prompts: Write C# code that uses Aspose.Cells to apply the UNIQUE function as a dynamic array formula and iterate over the spilled results. | Show how to refresh dynamic array formulas after modifying the source range in an Aspose.Cells workbook. | Explain how to programmatically determine the spill range size produced by a UNIQUE formula in Aspose.Cells.

using System;
using Aspose.Cells;

namespace UniqueDynamicArrayDemo
{
    // Demonstrates how to create a workbook, fill C1:C6 with duplicate fruit names, set the dynamic array formula "=UNIQUE(C1:C6)" in D1 using SetDynamicArrayFormula, calculate and refresh the spill range, read the unique values from column D, and save the file as UniqueDynamicArrayDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate column C (index 2) with duplicate values
            string[] sourceData = { "Apple", "Banana", "Apple", "Orange", "Banana", "Grape" };
            for (int i = 0; i < sourceData.Length; i++)
            {
                cells[i, 2].PutValue(sourceData[i]); // C1:C6
            }

            // Set a dynamic array formula in D1 that returns unique values from C1:C6
            Cell d1 = cells[0, 3]; // D1
            string uniqueFormula = "=UNIQUE(C1:C6)";
            d1.SetDynamicArrayFormula(uniqueFormula, new FormulaParseOptions(), true);

            // Calculate formulas and refresh dynamic array spill ranges
            wb.CalculateFormula();
            wb.RefreshDynamicArrayFormulas(true);

            // Output the spilled results from column D
            Console.WriteLine("Unique values spilled into column D:");
            for (int row = 0; row < cells.MaxDataRow + 1; row++)
            {
                Cell cell = cells[row, 3]; // D column
                if (cell != null && cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                {
                    Console.WriteLine($"{cell.Name}: {cell.StringValue}");
                }
            }

            // Save the workbook
            wb.Save("UniqueDynamicArrayDemo.xlsx");
        }
    }
}
