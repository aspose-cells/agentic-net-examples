using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            AsposeRange mainRange = cells.CreateRange("A1", "C3");
            mainRange.Name = "MainRange";

            int value = 1;
            for (int r = 0; r < mainRange.RowCount; r++)
            {
                for (int c = 0; c < mainRange.ColumnCount; c++)
                {
                    mainRange[r, c].PutValue(value++);
                }
            }

            cells.AddRange(mainRange);

            CellArea insertArea = CellArea.CreateCellArea(1, 0, 1, mainRange.ColumnCount - 1);
            cells.InsertRange(insertArea, ShiftType.Down);

            CellArea clearArea = CellArea.CreateCellArea("B2", "C3");
            cells.ClearRange(clearArea);

            workbook.Save("RangeDemo.xlsx");
        }
    }
}