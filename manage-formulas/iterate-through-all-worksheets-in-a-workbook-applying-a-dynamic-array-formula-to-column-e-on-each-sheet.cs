// Title: C# – Set a SEQUENCE dynamic array formula in column E of every worksheet using Aspose.Cells
// Description: Creates a workbook, adds a second sheet, iterates over all worksheets, inserts the dynamic‑array formula =SEQUENCE(5) into cell E1 with SetDynamicArrayFormula, forces immediate calculation, refreshes spill ranges, and saves the file as DynamicArrayResult.xlsx.
// Keywords: Aspose.Cells | C# | .NET | dynamic array formula | SetDynamicArrayFormula | SEQUENCE function | iterate worksheets | refresh spill range | calculate workbook formulas | multi‑sheet Excel automation
// Common Searches: Aspose.Cells set dynamic array formula on all sheets | C# loop through worksheets and apply SEQUENCE formula | Refresh dynamic array spill ranges Aspose.Cells | How to add a dynamic array to column E in each worksheet | Aspose.Cells SetDynamicArrayFormula example
// Developer Intent: Insert the same dynamic‑array formula into column E of every worksheet in a workbook and ensure the formulas are calculated and spill ranges are updated.
// Use Cases: Generate a numbered list (1‑5) in column E of each sheet for template‑based reporting. | Automatically populate multi‑sheet workbooks with calculated arrays without manual entry. | Maintain consistent dynamic‑array behavior across all worksheets after bulk updates.
// AI Prompts: Write C# code that loops through every worksheet in an Aspose.Cells workbook and sets =SEQUENCE(5) in cell E1, then calculates and refreshes the formulas. | Show how to replace the SEQUENCE formula with FILTER or SORT in column E for each worksheet using Aspose.Cells for .NET. | Explain the purpose of RefreshDynamicArrayFormulas after calling SetDynamicArrayFormula in a multi‑sheet workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayExample
{
    // Creates a workbook, adds a second sheet, iterates over all worksheets, inserts the dynamic‑array formula =SEQUENCE(5) into cell E1 with SetDynamicArrayFormula, forces immediate calculation, refreshes spill ranges, and saves the file as DynamicArrayResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Add a second worksheet to demonstrate iteration over multiple sheets
            workbook.Worksheets.Add("Sheet2");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the cell at the top of column E (E1) for the current sheet
                Cell targetCell = sheet.Cells["E1"];

                // Apply a dynamic array formula that will spill into the rows below.
                // Example formula: SEQUENCE(5) creates a vertical array of numbers 1‑5.
                // The FormulaParseOptions instance uses default settings.
                targetCell.SetDynamicArrayFormula(
                    "=SEQUENCE(5)",               // dynamic array formula
                    new FormulaParseOptions(),    // parsing options
                    true);                        // calculate the values immediately
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Refresh dynamic array formulas to ensure spill ranges are updated
            workbook.RefreshDynamicArrayFormulas(true);

            // Save the workbook (using the standard save rule)
            workbook.Save("DynamicArrayResult.xlsx");
        }
    }
}
