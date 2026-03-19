using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data in the worksheet (A1 to E10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // -------------------------------------------------
            // 1. Create a secondary range that will be used for
            //    data segmentation (e.g., B2:D5)
            // -------------------------------------------------
            AsposeRange secondaryRange = cells.CreateRange("B2", "D5");
            secondaryRange.Name = "SegmentRange";

            // Add the range to the worksheet's range collection
            cells.AddRange(secondaryRange);

            // -------------------------------------------------
            // 2. Apply a style to the secondary range to highlight it
            // -------------------------------------------------
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;
            style.Font.IsBold = true;
            secondaryRange.ApplyStyle(style, new StyleFlag
            {
                Font = true,
                CellShading = true
            });

            // -------------------------------------------------
            // 3. Clear the contents of a sub‑area inside the secondary range
            //    (e.g., clear C3:C4)
            // -------------------------------------------------
            CellArea clearArea = CellArea.CreateCellArea("C3", "C4");
            cells.ClearRange(clearArea);

            // -------------------------------------------------
            // 4. Insert a blank row above the secondary range and shift cells down
            // -------------------------------------------------
            int lastColumn = secondaryRange.FirstColumn + secondaryRange.ColumnCount - 1;
            CellArea insertArea = CellArea.CreateCellArea(
                secondaryRange.FirstRow,
                secondaryRange.FirstColumn,
                secondaryRange.FirstRow,
                lastColumn);
            // Insert one row (shiftNumber = 1) and shift cells down
            cells.InsertRange(insertArea, 1, ShiftType.Down);

            // -------------------------------------------------
            // 5. Copy the secondary range to a new location (e.g., G2:I5)
            // -------------------------------------------------
            AsposeRange destinationRange = cells.CreateRange("G2", "I5");
            destinationRange.Copy(secondaryRange);

            // -------------------------------------------------
            // 6. Save the workbook to an XLSX file
            // -------------------------------------------------
            workbook.Save("SecondaryRangeDemo.xlsx");
        }
    }
}