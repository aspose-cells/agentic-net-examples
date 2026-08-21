// Title: Recalculate All Formulas in an Aspose.Cells Workbook (C#)
// Description: Shows how to change cell values, invoke Workbook.CalculateFormula to recompute every dependent formula, read the results, and save the updated workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | Workbook.CalculateFormula | C# formula recalculation | recalculate dependent formulas | update cell values | calculate all formulas | Excel formula refresh .NET | Aspose.Cells example
// Common Searches: Aspose.Cells recalculate formulas after changing cells | Workbook.CalculateFormula C# example | How to refresh Excel formulas with Aspose.Cells | Recalculate all formulas in a workbook using Aspose.Cells | Aspose.Cells calculate all formulas programmatically
// Developer Intent: Refresh all workbook formulas after data changes.
// Use Cases: Refresh a financial model after bulk data import | Update engineering calculations when input parameters are modified | Regenerate a report with latest totals before exporting | Batch edit multiple worksheets and ensure formulas are up‑to‑date
// AI Prompts: Provide a C# snippet that modifies several cells and calls Workbook.CalculateFormula to recalculate the entire workbook. | Explain how to retrieve a cell's calculated value after invoking Workbook.CalculateFormula in Aspose.Cells. | Show how to recalculate formulas for a single worksheet instead of the whole workbook using Aspose.Cells. | Discuss performance tips for using Workbook.CalculateFormula on large workbooks.

using System;
using Aspose.Cells;

namespace RecalculateFormulasDemo
{
    // Shows how to change cell values, invoke Workbook.CalculateFormula to recompute every dependent formula, read the results, and save the updated workbook with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Use the default constructor as per the create rule

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some initial data and formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].Formula = "=A1*2";   // Depends on A1
            cells["B2"].Formula = "=A2*3";   // Depends on A2
            cells["C1"].Formula = "=B1+B2";  // Depends on B1 and B2

            // At this point formulas are not calculated yet
            Console.WriteLine("Before modification and calculation:");
            Console.WriteLine($"C1 (raw): {cells["C1"].StringValue}");

            // Modify worksheet data that affect formulas
            cells["A1"].PutValue(15); // Change value used by B1
            cells["A2"].PutValue(25); // Change value used by B2

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula(); // Uses the CalculateFormula() method as required

            // Display the updated results
            Console.WriteLine("\nAfter modification and calculation:");
            Console.WriteLine($"B1 result: {cells["B1"].IntValue}");
            Console.WriteLine($"B2 result: {cells["B2"].IntValue}");
            Console.WriteLine($"C1 result: {cells["C1"].IntValue}");

            // Save the workbook to verify persisted values (save rule)
            workbook.Save("RecalculatedFormulas.xlsx");
        }
    }
}
