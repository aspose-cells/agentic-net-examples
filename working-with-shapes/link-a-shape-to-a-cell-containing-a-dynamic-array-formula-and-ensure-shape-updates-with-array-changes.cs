using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDynamicArrayShapeDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 2. Prepare data that will drive the dynamic array formula
            //    Cell B1 will hold the number of items to generate
            // ------------------------------------------------------------
            cells["B1"].PutValue(3); // initial size = 3

            // ------------------------------------------------------------
            // 3. Set a dynamic array formula in A1 that spills into A1:A{n}
            //    Formula: =SEQUENCE(B1)  (creates a vertical list 1..B1)
            // ------------------------------------------------------------
            cells["A1"].SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // ------------------------------------------------------------
            // 4. Add a ListBox shape that will display the spilled values.
            //    The shape is placed at row 5, column 1 with width/height.
            // ------------------------------------------------------------
            ListBox listBox = (ListBox)sheet.Shapes.AddListBox(5, 1, 120, 80, 1, 5);

            // ------------------------------------------------------------
            // 5. Link the ListBox to the first cell of the spill range (A1).
            //    The shape will read the values from the spill range automatically.
            // ------------------------------------------------------------
            listBox.SetLinkedCell("$A$1", false, true);

            // ------------------------------------------------------------
            // 6. Populate the ListBox's input range with the spill range.
            //    Use the spill operator (#) to refer to the whole dynamic array.
            // ------------------------------------------------------------
            listBox.SetInputRange("A1#", false, false);

            // ------------------------------------------------------------
            // 7. Update the shape's selected value based on the linked cell.
            // ------------------------------------------------------------
            sheet.Shapes.UpdateSelectedValue();

            // ------------------------------------------------------------
            // 8. Change the driving parameter (B1) to expand the array.
            // ------------------------------------------------------------
            cells["B1"].PutValue(6); // now the array should spill 1..6

            // ------------------------------------------------------------
            // 9. Refresh dynamic array formulas so the spill range is recalculated.
            //    Pass true to also calculate the new values.
            // ------------------------------------------------------------
            workbook.RefreshDynamicArrayFormulas(true);

            // ------------------------------------------------------------
            // 10. After refreshing, update the shape again so it reflects the new spill.
            // ------------------------------------------------------------
            sheet.Shapes.UpdateSelectedValue();

            // ------------------------------------------------------------
            // 11. (Optional) Verify the spill range values in the console.
            // ------------------------------------------------------------
            Console.WriteLine("Spill values after refresh:");
            for (int row = 0; row < 6; row++)
            {
                Console.WriteLine($"A{row + 1} = {cells[row, 0].Value}");
            }

            // ------------------------------------------------------------
            // 12. Save the workbook to a file.
            // ------------------------------------------------------------
            workbook.Save("DynamicArrayLinkedShape.xlsx");
        }
    }
}