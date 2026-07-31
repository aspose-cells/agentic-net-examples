// Title: Aspose.Cells for .NET – SEQUENCE dynamic array spill, insert rows inside the spill, and refresh the range
// Description: Demonstrates how to set a =SEQUENCE dynamic array formula that spills vertically, insert a row within the spilled area, recalculate with RefreshDynamicArrayFormulas, and save the workbook using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | .NET | dynamic array formula | SEQUENCE function | spill range | InsertRows | RefreshDynamicArrayFormulas | Excel automation | spreadsheet recalculation
// Common Searches: Aspose.Cells set dynamic array formula SEQUENCE | how to insert row inside spilled array Aspose.Cells | refresh dynamic array after row insertion C# | Aspose.Cells spill range update | C# example dynamic array spill and refresh
// Developer Intent: Create a SEQUENCE‑driven dynamic array, add a row inside its spill area, and refresh the formula so the spill adjusts automatically.
// Use Cases: Generate a vertical list whose length is controlled by a count cell. | Add custom entries within a spilled list without breaking the underlying formula. | Maintain correct spill dimensions after structural changes such as row insertion or deletion.
// AI Prompts: Provide C# code that uses SetDynamicArrayFormula("=SEQUENCE(B1)") and then inserts a row while preserving the formula reference in Aspose.Cells. | Explain the role of RefreshDynamicArrayFormulas after inserting rows inside a spilled dynamic array range. | Show how to update a SEQUENCE‑based spill when the source count cell changes and rows are added or removed.

using System;
using Aspose.Cells;

namespace DynamicArraySpillDemo
{
    // Demonstrates how to set a =SEQUENCE dynamic array formula that spills vertically, insert a row within the spilled area, recalculate with RefreshDynamicArrayFormulas, and save the workbook using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Prepare data that the dynamic array formula will use.
            //    Cell B1 will hold the number of rows the SEQUENCE function should generate.
            // ------------------------------------------------------------
            cells["B1"].PutValue(3); // Initial spill size = 3 rows

            // ------------------------------------------------------------
            // 2. Set a dynamic array formula that spills into empty rows.
            //    The formula is placed in A2 and will spill downwards based on B1.
            // ------------------------------------------------------------
            Cell startCell = cells["A2"];
            // The formula will generate a vertical sequence: 1,2,3,...
            startCell.SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // Calculate formulas so the spill range is populated
            workbook.CalculateFormula();

            // Display the initial spilled values
            Console.WriteLine("Initial spill (A2:A4):");
            for (int i = 2; i <= 4; i++)
            {
                Console.WriteLine($"A{i} = {cells[i - 1, 0].Value}");
            }

            // ------------------------------------------------------------
            // 3. Insert a new row within the spilled range to shift the existing data down.
            //    Insert at row index 2 (which corresponds to Excel row 3).
            // ------------------------------------------------------------
            // Insert 1 row, update references so the formula still points to B1
            cells.InsertRows(2, 1, true);

            // Optionally put a value in the newly inserted row to see the shift effect
            cells["A3"].PutValue("Inserted Row");

            // ------------------------------------------------------------
            // 4. Refresh dynamic array formulas so the spill range adapts to the new layout.
            //    The 'calculate' flag is true to recalculate the formula values.
            // ------------------------------------------------------------
            workbook.RefreshDynamicArrayFormulas(true);

            // Display the spilled values after insertion and refresh
            Console.WriteLine("\nSpill after inserting a row and refreshing:");
            // Determine the new spill range (A2 downwards, length still based on B1 = 3)
            for (int i = 2; i <= 5; i++) // Expect rows A2:A5 now
            {
                Console.WriteLine($"A{i} = {cells[i - 1, 0].Value}");
            }

            // ------------------------------------------------------------
            // 5. Save the workbook to verify the result in Excel.
            // ------------------------------------------------------------
            workbook.Save("DynamicArraySpillShift.xlsx");
        }
    }
}
