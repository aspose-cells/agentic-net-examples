// Title: Refresh a FILTER dynamic array formula after changing source data with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply a FILTER dynamic array formula using SetDynamicArrayFormula, modify the source range, invoke RefreshDynamicArrayFormulas, recalculate the workbook, read the updated spill values, and save the file.
// Keywords: Aspose.Cells FILTER dynamic array | SetDynamicArrayFormula .NET | RefreshDynamicArrayFormulas | update source data spreadsheet | recalculate dynamic array formula | C# Aspose.Cells example | spill range handling
// Common Searches: Aspose.Cells change source data for FILTER formula | Refresh dynamic array after cell update Aspose.Cells | C# set and recalculate FILTER dynamic array | How to refresh spilled results in Aspose.Cells | Update FILTER criteria programmatically .NET
// Developer Intent: Update the source cells that a FILTER dynamic array depends on and refresh the formula to obtain the new spill results.
// Use Cases: Generate reports where FILTER results automatically reflect data edits. | Build dashboards that need real‑time recalculation after user input. | Programmatically adjust spreadsheet calculations before exporting or sharing.
// AI Prompts: Write C# code with Aspose.Cells to set a FILTER dynamic array, modify a source cell, and refresh the formula to get updated spill values. | Explain why RefreshDynamicArrayFormulas must be called after changing source data for dynamic arrays in Aspose.Cells. | Show how to iterate over the spill range of a FILTER formula and print each value using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply a FILTER dynamic array formula using SetDynamicArrayFormula, modify the source range, invoke RefreshDynamicArrayFormulas, recalculate the workbook, read the updated spill values, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate source data for the FILTER formula
        cells["A2"].PutValue(10);
        cells["A3"].PutValue(20);
        cells["A4"].PutValue(30);
        cells["A5"].PutValue(40);

        cells["B2"].PutValue(5);
        cells["B3"].PutValue(15);
        cells["B4"].PutValue(25);
        cells["B5"].PutValue(35);

        // Set a FILTER dynamic array formula in C2:
        // =FILTER(A2:A5, B2:B5>20)
        Cell formulaCell = cells["C2"];
        string formula = "=FILTER(A2:A5, B2:B5>20)";
        formulaCell.SetDynamicArrayFormula(formula, new FormulaParseOptions(), true);

        // Calculate initial results so the spill range is populated
        wb.CalculateFormula();

        // Change source data that influences the FILTER result
        // For example, make B4 greater than 20 so its corresponding A4 should appear in the result
        cells["B4"].PutValue(50);

        // Refresh dynamic array formulas and recalculate their values
        wb.RefreshDynamicArrayFormulas(true);

        // Output the values from the spill range (C2:C5)
        for (int row = 2; row <= 5; row++)
        {
            Console.WriteLine($"C{row} = {cells[$"C{row}"].Value}");
        }

        // Save the workbook with the updated dynamic array results
        wb.Save("FilterDynamicArrayDemo.xlsx");
    }
}
