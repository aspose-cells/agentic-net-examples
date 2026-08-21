// Title: Refresh Dynamic Array Formulas After Changing Source Cells with Workbook.CalculateFormula (C# Aspose.Cells)
// Description: Demonstrates how to create a workbook, insert source values, set a dynamic‑array SEQUENCE formula, calculate all formulas, modify the source data, invoke Workbook.CalculateFormula to refresh the spilled results across every worksheet, display the updated values, and save the file.
// Keywords: Aspose.Cells | .NET | C# | dynamic array | SEQUENCE formula | Workbook.CalculateFormula | recalculate formulas | spilled array refresh | update source cells | multiple worksheets
// Common Searches: Aspose.Cells recalculate dynamic array after data change | Workbook.CalculateFormula refresh spilled array C# | How to update SEQUENCE formula in Aspose.Cells | Refresh all formulas in workbook Aspose.Cells .NET | Dynamic array formula recalc across sheets Aspose
// Developer Intent: Recalculate every dynamic‑array formula after programmatically changing the source cells.
// Use Cases: Refresh dynamic‑array results after bulk data import before exporting the workbook. | Ensure formulas on multiple sheets stay in sync after batch updates to source ranges. | Generate up‑to‑date SEQUENCE outputs for financial or reporting dashboards.
// AI Prompts: Provide C# code that modifies source cells and calls wb.CalculateFormula to refresh dynamic‑array formulas in all worksheets using Aspose.Cells. | Explain how Workbook.CalculateFormula interacts with dynamic‑array (spilled) formulas and how to guarantee they recalculate after cell edits. | Show a pattern for looping through several worksheets and invoking CalculateFormula after batch updates in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, insert source values, set a dynamic‑array SEQUENCE formula, calculate all formulas, modify the source data, invoke Workbook.CalculateFormula to refresh the spilled results across every worksheet, display the updated values, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate source data that will be used by a dynamic array formula
        cells["A1"].PutValue(1);
        cells["A2"].PutValue(2);
        cells["A3"].PutValue(3);

        // Set a dynamic array formula in B1 that spills into B1:B3
        // The formula generates a sequence of numbers; calculateValue = true ensures initial calculation
        cells["B1"].SetDynamicArrayFormula("=SEQUENCE(3)", new FormulaParseOptions(), true);

        // Initial calculation of all formulas, including the dynamic array formula
        wb.CalculateFormula();

        // Modify the source data that influences the dynamic array formula
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);

        // Recalculate all formulas across all worksheets.
        // This refreshes the dynamic array formulas to reflect the updated data.
        wb.CalculateFormula();

        // Output the refreshed dynamic array results
        Console.WriteLine("Dynamic array results after data change:");
        for (int i = 0; i < 3; i++)
        {
            Console.WriteLine($"B{i + 1}: {cells[i, 1].Value}");
        }

        // Save the workbook (saving rule)
        wb.Save("DynamicArrayRefreshDemo.xlsx");
    }
}
