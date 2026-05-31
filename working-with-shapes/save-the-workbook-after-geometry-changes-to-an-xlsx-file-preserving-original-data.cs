using System;
using Aspose.Cells;

namespace AsposeCellsGeometrySaveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (preserves original data)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Geometry changes -----

            // 1. Insert a new row at index 2 (third row), shifting existing rows down
            CellArea insertArea = CellArea.CreateCellArea(2, 0, 2, sheet.Cells.MaxColumn);
            sheet.Cells.InsertRange(insertArea, 1, ShiftType.Down, true);

            // 2. Set column width for column B (index 1) to 20 characters
            sheet.Cells.SetColumnWidth(1, 20);

            // 3. Merge cells A1 to C1 (first row, columns 0-2)
            sheet.Cells.Merge(0, 0, 1, 3);

            // Save the modified workbook to a new XLSX file, preserving all data
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}