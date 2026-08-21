// Title: Link a ListBox shape to a dynamic array formula and auto‑refresh with Aspose.Cells for .NET
// Description: Shows how to create a workbook, apply a SEQUENCE dynamic‑array formula, add a ListBox shape, link it to the spill range’s top‑left cell, recalculate the array when the driver cell changes, and synchronize the shape using RefreshDynamicArrayFormulas and Shapes.UpdateSelectedValue.
// Keywords: Aspose.Cells | C# | .NET | dynamic array | SEQUENCE formula | shape linking | ListBox | LinkedCell | RefreshDynamicArrayFormulas | UpdateSelectedValue | spill range | Excel automation
// Common Searches: Aspose.Cells link shape to dynamic array | update ListBox after SEQUENCE formula changes | RefreshDynamicArrayFormulas example C# | sync shape with spilled array Aspose.Cells | C# Aspose.Cells linked cell shape
// Developer Intent: Demonstrate how to bind a shape to a cell containing a dynamic‑array formula and keep the shape synchronized when the array size changes.
// Use Cases: Create an interactive ListBox that reflects a calculated list whose length can grow or shrink based on input data. | Generate server‑side reports where shapes automatically update after formula recalculation. | Build dashboards that adjust visual controls in real time as underlying spreadsheet data changes.
// AI Prompts: Write C# code that links a ListBox shape to the first cell of a SEQUENCE dynamic‑array spill range and updates the shape after the array expands using Aspose.Cells. | Show how to refresh dynamic‑array formulas and synchronize linked shapes in an Aspose.Cells workbook. | Provide an example of programmatically changing the driver cell, calling RefreshDynamicArrayFormulas, and then invoking Shapes.UpdateSelectedValue to keep a shape in sync.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDynamicArrayShapeLink
{
    // Shows how to create a workbook, apply a SEQUENCE dynamic‑array formula, add a ListBox shape, link it to the spill range’s top‑left cell, recalculate the array when the driver cell changes, and synchronize the shape using RefreshDynamicArrayFormulas and Shapes.UpdateSelectedValue.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Prepare data that will drive the dynamic array formula
                //    Column B will contain the size of the sequence
                cells["B1"].PutValue(3); // initial size = 3

                // 3. Set a dynamic array formula in A1 that spills vertically.
                //    The formula =SEQUENCE(B1) will produce a vertical list of numbers 1..B1
                cells["A1"].SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

                // 4. Add a ListBox shape (any shape that supports LinkedCell)
                //    Parameters: upper left row, upper left column, upper left offset X, offset Y,
                //                width, height
                ListBox listBox = (ListBox)sheet.Shapes.AddListBox(2, 2, 0, 0, 120, 80);

                // 5. Link the shape to the top‑left cell of the spill range (A1)
                //    SetLinkedCell(string formula, bool isR1C1, bool isLocal)
                listBox.SetLinkedCell("$A$1", false, true);

                // 6. Populate the ListBox input range (optional, just for demonstration)
                //    Here we use the same spill range as the input source.
                listBox.SetInputRange("A1:A3", false, false);
                listBox.SelectedIndex = 0; // select first item

                // 7. Ensure the shape reflects the current linked cell value
                sheet.Shapes.UpdateSelectedValue();

                // 8. Change the driving data so the dynamic array expands
                cells["B1"].PutValue(5); // now the sequence should be 1..5

                // 9. Refresh dynamic array formulas so the spill range updates
                //    The first parameter 'true' tells Aspose.Cells to recalculate the formulas.
                workbook.RefreshDynamicArrayFormulas(true);

                // 10. After the spill range has changed, update the shape again
                sheet.Shapes.UpdateSelectedValue();

                // 11. Save the workbook to verify the result
                string outputPath = "DynamicArrayShapeLink.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
