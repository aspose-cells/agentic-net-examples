// Title: C# – Insert SORT Dynamic Array Formula in B2 and Validate Spill Range with Aspose.Cells
// Description: Shows how to build a workbook, fill A1:A5 with descending numbers, apply the Excel SORT dynamic‑array formula to cell B2 via SetDynamicArrayFormula, recalculate, refresh dynamic‑array formulas, obtain the spilled area using GetArrayRange, display the sorted values, and save the result as DynamicArraySortDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | dynamic array formula | SORT function | SetDynamicArrayFormula | GetArrayRange | RefreshDynamicArrayFormulas | spill range | Excel dynamic arrays
// Common Searches: Aspose.Cells set SORT formula C# | Get spill range of dynamic array Aspose.Cells | Refresh dynamic array formulas Aspose.Cells .NET | C# example for Excel SORT dynamic array | How to use SetDynamicArrayFormula Aspose.Cells
// Developer Intent: Add a SORT dynamic‑array formula to B2 and confirm that the resulting spill range expands correctly.
// Use Cases: Automatically sort a column of data inside a generated workbook. | Programmatically retrieve the extent of a spilled dynamic‑array result for further processing. | Ensure dynamic‑array calculations are up‑to‑date after workbook modifications. | Export verified sorted data to external systems.
// AI Prompts: Write C# code using Aspose.Cells to place =SORT(A1:A5) in B2, refresh formulas, and print the spill range. | Explain how GetArrayRange works with Excel dynamic arrays in Aspose.Cells and show a sample output. | Provide a step‑by‑step guide that verifies the spill area of a SORT dynamic array after workbook calculation.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArraySortDemo
{
    // Shows how to build a workbook, fill A1:A5 with descending numbers, apply the Excel SORT dynamic‑array formula to cell B2 via SetDynamicArrayFormula, recalculate, refresh dynamic‑array formulas, obtain the spilled area using GetArrayRange, display the sorted values, and save the result as DynamicArraySortDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(5 - i); // Values: 5,4,3,2,1
            }

            // Set a SORT dynamic array formula in cell B2
            Cell targetCell = cells["B2"];
            targetCell.SetDynamicArrayFormula("=SORT(A1:A5)", new FormulaParseOptions(), true);

            // Calculate formulas and refresh dynamic array formulas to ensure spill range is updated
            workbook.CalculateFormula();
            workbook.RefreshDynamicArrayFormulas(true);

            // Retrieve the spilled range of the dynamic array formula
            CellArea spillRange = targetCell.GetArrayRange();

            // Output spilled range information
            Console.WriteLine($"Spill range start: Row {spillRange.StartRow + 1}, Column {spillRange.StartColumn + 1}");
            Console.WriteLine($"Spill range end:   Row {spillRange.EndRow + 1}, Column {spillRange.EndColumn + 1}");

            // Print the values inside the spilled range to verify correct sorting
            Console.WriteLine("Sorted values in spilled range:");
            for (int row = spillRange.StartRow; row <= spillRange.EndRow; row++)
            {
                for (int col = spillRange.StartColumn; col <= spillRange.EndColumn; col++)
                {
                    Console.Write(cells[row, col].Value + "\t");
                }
                Console.WriteLine();
            }

            // Save the workbook
            workbook.Save("DynamicArraySortDemo.xlsx");
        }
    }
}
