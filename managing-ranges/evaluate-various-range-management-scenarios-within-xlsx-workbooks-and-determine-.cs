using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeManagementDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate some sample data (A1:D5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // 3. Create a range using address string ("B2:C4")
            AsposeRange addressRange = cells.CreateRange("B2", "C4");
            addressRange.Name = "AddressRange";

            // 4. Clear only the contents of the address range (keep formatting)
            cells.ClearContents(addressRange.FirstRow, addressRange.FirstColumn,
                                addressRange.RowCount, addressRange.ColumnCount);

            // 5. Create a range using numeric indices (A1:B3)
            AsposeRange indexRange = cells.CreateRange(0, 0, 3, 2);
            indexRange.Name = "IndexRange";

            // 6. Apply a background style to the index range
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            indexRange.ApplyStyle(style, new StyleFlag { CellShading = true });

            // 7. Delete a sub‑range (C3:D5) and shift cells left/up
            int delFirstRow = 2;   // C3 (zero‑based)
            int delFirstCol = 2;   // C3 (zero‑based)
            int delRows = 3;       // rows C3:D5 => 3 rows
            int delCols = 2;       // columns C and D => 2 columns
            cells.DeleteRange(delFirstRow, delFirstCol, delRows, delCols, ShiftType.Left);

            // 8. Insert a blank range at E2 (shift existing cells down)
            CellArea insertArea = new CellArea
            {
                StartRow = 1,
                EndRow = 1,
                StartColumn = 4,
                EndColumn = 4
            };
            cells.InsertRange(insertArea, ShiftType.Down);

            // 9. Move a range (A1:B2) to start at G5
            CellArea moveSource = new CellArea
            {
                StartRow = 0,
                EndRow = 1,
                StartColumn = 0,
                EndColumn = 1
            };
            cells.MoveRange(moveSource, 4, 6); // target row 4 (row 5), column 6 (column G)

            // 10. Merge a block of cells (F1:G3) into a single cell
            cells.Merge(0, 5, 3, 2); // rows 0‑2 (3 rows), columns 5‑6 (2 columns)

            // 11. Unmerge the previously merged block
            cells.UnMerge(0, 5, 3, 2);

            // 12. Clear all merged cells in the worksheet (if any remain)
            cells.ClearMergedCells();

            // 13. Delete blank rows and columns in the used range
            cells.DeleteBlankRows();
            cells.DeleteBlankColumns();

            // 14. Save the workbook to disk
            workbook.Save("RangeManagementDemo.xlsx");
        }
    }
}