// Title: Aspose.Cells for .NET: Dynamically Link a Shape to a VLOOKUP Result Cell
// Description: C# example that creates a lookup table, adds a VLOOKUP formula, inserts a rectangle shape, links the shape to the formula cell with SetLinkedCell, recalculates the workbook, and refreshes the shape text using UpdateSelectedValue so the shape always reflects the current lookup result.
// Keywords: Aspose.Cells | C# | .NET | shape linked cell | SetLinkedCell | UpdateSelectedValue | VLOOKUP | dynamic shape text | Excel automation | worksheet shape binding
// Common Searches: Aspose.Cells link shape to cell example | SetLinkedCell method with VLOOKUP result C# | Refresh shape text after formula change Aspose.Cells | How to bind a rectangle to a cell in Aspose.Cells .NET | Update linked shape when lookup key changes
// Developer Intent: Bind a worksheet shape to a cell that contains a VLOOKUP formula and keep the shape’s displayed value synchronized with formula updates.
// Use Cases: Display product prices inside shapes on a sales dashboard that automatically update when the selected item changes. | Create financial reports where key metrics are shown in shapes and stay current after data edits. | Build interactive Excel templates with shapes that reflect lookup‑driven calculations without manual refresh.
// AI Prompts: Show me C# code that links a rectangle shape to a VLOOKUP result cell using Aspose.Cells and updates the shape after changing the lookup key. | Explain how SetLinkedCell and UpdateSelectedValue work together to keep a shape synchronized with a formula in Aspose.Cells for .NET. | Provide a step‑by‑step guide for handling multiple shapes, each linked to different VLOOKUP results, in the same worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkedCellDemo
{
    // C# example that creates a lookup table, adds a VLOOKUP formula, inserts a rectangle shape, links the shape to the formula cell with SetLinkedCell, recalculates the workbook, and refreshes the shape text using UpdateSelectedValue so the shape always reflects the current lookup result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate lookup table (A1:B5)
            // -------------------------------------------------
            sheet.Cells["A1"].Value = "Item";
            sheet.Cells["B1"].Value = "Price";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 1.2;
            sheet.Cells["A3"].Value = "Banana";
            sheet.Cells["B3"].Value = 0.8;
            sheet.Cells["A4"].Value = "Cherry";
            sheet.Cells["B4"].Value = 2.5;
            sheet.Cells["A5"].Value = "Date";
            sheet.Cells["B5"].Value = 3.0;

            // -------------------------------------------------
            // Cell C1 will hold the lookup key (e.g., "Banana")
            // -------------------------------------------------
            sheet.Cells["C1"].Value = "Banana";

            // -------------------------------------------------
            // D1 contains the VLOOKUP formula that returns the price
            // =VLOOKUP(C1, $A$2:$B$5, 2, FALSE)
            // -------------------------------------------------
            sheet.Cells["D1"].Formula = "=VLOOKUP(C1, $A$2:$B$5, 2, FALSE)";

            // -------------------------------------------------
            // Add a rectangle shape that will display the lookup result
            // -------------------------------------------------
            // Parameters: upper left row, upper left column, upper left offset (pixels),
            // lower right row, lower right column, lower right offset (pixels)
            Shape rect = sheet.Shapes.AddRectangle(2, 2, 0, 4, 2, 0);
            rect.Text = "Lookup Result";

            // Link the shape to the cell containing the VLOOKUP result (D1)
            // Using SetLinkedCell method (formula, isR1C1, isLocal)
            rect.SetLinkedCell("$D$1", false, true);

            // -------------------------------------------------
            // Recalculate the workbook so the formula evaluates
            // -------------------------------------------------
            workbook.CalculateFormula();

            // Update the shape's displayed value based on the linked cell
            rect.UpdateSelectedValue();

            // -------------------------------------------------
            // Change the lookup key to demonstrate dynamic update
            // -------------------------------------------------
            sheet.Cells["C1"].Value = "Cherry";

            // Recalculate and refresh the shape again
            workbook.CalculateFormula();
            rect.UpdateSelectedValue();

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ShapeLinkedCellVLookupDemo.xlsx");
        }
    }
}
