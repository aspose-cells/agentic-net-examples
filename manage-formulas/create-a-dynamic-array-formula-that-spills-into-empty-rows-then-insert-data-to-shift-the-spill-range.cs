// Title: Set and Shift a SEQUENCE Dynamic Array Formula with Row Insertion Using Aspose.Cells for .NET
// Description: Shows how to create a workbook, put a control value in B1, assign the dynamic array formula =SEQUENCE(B1) to A1, calculate and refresh the spill, insert rows to move the spilled range, change B1 to enlarge the array, refresh again, and save the file. Illustrates programmatic handling of dynamic‑array spills and their repositioning in C#.
// Keywords: Aspose.Cells | dynamic array | SEQUENCE formula | spill range | insert rows | refresh formulas | C# | .NET | Workbook manipulation | CalculateFormula | RefreshDynamicArrayFormulas
// Common Searches: How to create a SEQUENCE dynamic array that spills into rows with Aspose.Cells | How to shift a spilled dynamic array by inserting rows in Aspose.Cells for .NET | How to expand a SEQUENCE array after row insertion using Aspose.Cells | Refresh dynamic array formulas after changing a control cell in Aspose.Cells
// Developer Intent: Create a SEQUENCE dynamic array that spills vertically, insert rows to shift the spilled range, then enlarge the array by updating the control cell and refreshing formulas.
// Use Cases: Generate a numbered list that automatically grows when the control cell value changes. | Insert rows into a report while keeping a spilled dynamic array correctly positioned. | Adjust the size of a dynamic array after data insertion without rewriting the formula.
// AI Prompts: Write C# code with Aspose.Cells to set =SEQUENCE(B1) in A1, insert rows to move the spill, update B1, and refresh the formulas. | Provide an example that expands a spilled SEQUENCE array after inserting rows by changing the control cell and calling RefreshDynamicArrayFormulas. | Explain why CalculateFormula and RefreshDynamicArrayFormulas are needed when managing dynamic array spills in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayDemo
{
    // Shows how to create a workbook, put a control value in B1, assign the dynamic array formula =SEQUENCE(B1) to A1, calculate and refresh the spill, insert rows to move the spilled range, change B1 to enlarge the array, refresh again, and save the file. Illustrates programmatic handling of dynamic‑array spills and their repositioning in C#.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Put a value that will control the size of the dynamic array
            //    B1 = 3 means the SEQUENCE formula will spill into 3 rows (A1:A3)
            cells["B1"].PutValue(3);

            // 3. Set a dynamic array formula in A1 that spills based on B1
            //    The formula will generate a vertical sequence of numbers
            Cell startCell = cells["A1"];
            startCell.SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // 4. Calculate formulas and refresh dynamic array formulas so the spill occurs
            workbook.CalculateFormula();
            workbook.RefreshDynamicArrayFormulas(true);

            // 5. Insert two empty rows at row index 1 (i.e., between rows 1 and 2)
            //    This shifts the existing spilled range (A1:A3) down by two rows
            cells.InsertRows(1, 2, true); // rowIndex = 1 (second row), totalRows = 2

            // 6. Update the controlling value to expand the spill range
            //    Now B1 = 5, so the formula should spill into five rows (A1:A5)
            cells["B1"].PutValue(5);

            // 7. Refresh dynamic array formulas again to recalculate the spill area
            workbook.RefreshDynamicArrayFormulas(true);

            // 8. Save the workbook to verify the result
            workbook.Save("DynamicArraySpillShiftDemo.xlsx");
        }
    }
}
