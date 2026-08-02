// Title: Link a ListBox shape to a dynamic array spill and auto‑sync with Aspose.Cells for .NET
// Description: Creates a workbook, sets B1 as the size for a SEQUENCE dynamic array in A1, adds a ListBox shape, links its selected cell and input range to the spill, refreshes formulas, updates the shape, changes B1 to expand the array, and saves the file using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | dynamic array formula | SEQUENCE function | ListBox shape | shape linking | SetLinkedCell | SetInputRange | RefreshDynamicArrayFormulas | UpdateSelectedValue | spreadsheet automation
// Common Searches: Aspose.Cells link shape to dynamic array spill | C# update ListBox after SEQUENCE formula changes | RefreshDynamicArrayFormulas example | SetLinkedCell for ListBox in Aspose.Cells | Auto‑expand ListBox items with dynamic array
// Developer Intent: Bind a ListBox shape to a dynamic array spill and keep it synchronized when the array size changes.
// Use Cases: Show SEQUENCE results in a ListBox that grows or shrinks automatically. | Maintain the selected value of a shape in sync with the first element of a spill range. | Generate workbooks where form controls reflect the latest calculated data without manual edits.
// AI Prompts: Generate C# code that links a ListBox shape to a SEQUENCE dynamic array spill and updates it after the source cell changes using Aspose.Cells. | Explain how to use RefreshDynamicArrayFormulas and UpdateSelectedValue to keep a shape synchronized with a dynamic array in Aspose.Cells for .NET. | Provide step‑by‑step instructions to bind a shape’s linked cell and input range to a dynamic array formula in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDynamicArrayShapeDemo
{
    // Creates a workbook, sets B1 as the size for a SEQUENCE dynamic array in A1, adds a ListBox shape, links its selected cell and input range to the spill, refreshes formulas, updates the shape, changes B1 to expand the array, and saves the file using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // ------------------------------------------------------------
            // 1. Prepare data that drives a dynamic array formula
            // ------------------------------------------------------------
            // Cell B1 will hold the size of the sequence
            cells["B1"].PutValue(3);

            // Set a dynamic array formula in A1 that spills into A1:A3
            // The formula uses SEQUENCE which returns a vertical array
            cells["A1"].SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // ------------------------------------------------------------
            // 2. Add a ListBox shape and link it to the spilled range
            // ------------------------------------------------------------
            // Add a ListBox shape (type ListBox inherits from Shape)
            ListBox listBox = (ListBox)ws.Shapes.AddListBox(0, 0, 150, 100, 3, 20);

            // The spilled range of the dynamic array formula is A1:A3.
            // Link the ListBox's selected value to the first cell of the spill range (A1).
            // When the spill changes size, we will refresh the shape's selected value.
            listBox.SetLinkedCell("$A$1", false, true);

            // Optionally set the input range for the ListBox items to the spill range.
            // This makes the ListBox display the array values.
            listBox.SetInputRange("A1:A3", false, false);

            // ------------------------------------------------------------
            // 3. Initial calculation and shape synchronization
            // ------------------------------------------------------------
            // Calculate formulas and refresh dynamic array spill
            wb.CalculateFormula();
            wb.RefreshDynamicArrayFormulas(true);

            // Update the shape's selected value based on the linked cell
            ws.Shapes.UpdateSelectedValue();

            // ------------------------------------------------------------
            // 4. Change the driving data, refresh, and sync shape again
            // ------------------------------------------------------------
            // Change B1 to expand the sequence to 5 items
            cells["B1"].PutValue(5);

            // Refresh dynamic array formulas (re‑spill) and recalculate
            wb.RefreshDynamicArrayFormulas(true);
            wb.CalculateFormula();

            // Update the ListBox so it reflects the new spill range
            ws.Shapes.UpdateSelectedValue();

            // ------------------------------------------------------------
            // 5. Save the workbook (lifecycle rule: save)
            // ------------------------------------------------------------
            wb.Save("DynamicArrayLinkedShape.xlsx");
        }
    }
}
