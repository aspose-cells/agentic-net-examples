// Title: Update a linked ListBox shape after changing a source cell formula with Aspose.Cells for .NET
// Description: Demonstrates how to refresh a ListBox shape linked to a cell after a source cell formula is modified. The example sets an input range, links the shape to a cell, recalculates formulas, updates the linked cell value, and calls UpdateSelectedValue to keep the shape in sync before saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | ListBox shape | linked cell | UpdateSelectedValue | SetLinkedCell | SetInputRange | formula recalculation | Excel shape synchronization | Refresh linked shape
// Common Searches: Aspose.Cells refresh ListBox after formula change | Update linked shape after recalculating workbook | C# Aspose.Cells UpdateSelectedValue example | How to sync ListBox with calculated cell in Aspose.Cells | SetLinkedCell and UpdateSelectedValue usage
// Developer Intent: Synchronize a ListBox shape’s selected item with a new value after the source cell’s formula is changed and the workbook is recalculated.
// Use Cases: Keep form controls in generated Excel reports aligned with dynamic calculations. | Maintain consistency between dashboard shapes and underlying formula‑driven data. | Automate the refresh of linked shapes when source values are updated programmatically.
// AI Prompts: Generate C# code that updates a ListBox shape after a source cell formula is changed using Aspose.Cells. | Show how to use SetLinkedCell and UpdateSelectedValue to keep a shape synchronized with a calculated cell value. | Explain the steps to refresh a linked shape when its input range contains formula results.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to refresh a ListBox shape linked to a cell after a source cell formula is modified. The example sets an input range, links the shape to a cell, recalculates formulas, updates the linked cell value, and calls UpdateSelectedValue to keep the shape in sync before saving the workbook.
class UpdateLinkedShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some source data that will be used by the shape
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["A2"].Value = 20;
        sheet.Cells["A3"].Value = 30;

        // Add a ListBox shape to the worksheet
        Shape listBoxShape = sheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

        // Set the range that provides the list items for the ListBox
        listBoxShape.SetInputRange("$A$1:$A$3", false, false);

        // Link the selected value of the ListBox to cell B1
        listBoxShape.SetLinkedCell("$B$1", false, true);

        // Set an initial value in the linked cell (select the second item, value 20)
        sheet.Cells["B1"].Value = 20;

        // Update the shape so that its selected item reflects the linked cell value
        listBoxShape.UpdateSelectedValue();

        // Change the source data by assigning a formula to A2
        sheet.Cells["A2"].Formula = "=A1*3"; // A2 will become 30 after calculation

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Update the linked cell to the new value (select the third item, value 30)
        sheet.Cells["B1"].Value = 30;

        // Refresh the shape selection after the source change
        listBoxShape.UpdateSelectedValue();

        // Save the workbook with the updated shape
        workbook.Save("UpdatedLinkedShape.xlsx");
    }
}
