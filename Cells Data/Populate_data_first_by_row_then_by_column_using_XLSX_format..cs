using System;
using Aspose.Cells;

namespace AsposeCellsRowColumnDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // -------- Populate data row‑wise (horizontal) --------
            // Sample data that will be placed in the first row starting at cell A1
            string[] rowData = new string[] { "Row1_Col1", "Row1_Col2", "Row1_Col3", "Row1_Col4" };
            // isVertical = false => import horizontally (by row)
            cells.ImportArray(rowData, firstRow: 0, firstColumn: 0, isVertical: false);

            // -------- Populate data column‑wise (vertical) --------
            // Sample data that will be placed in the first column of a new block starting at cell D1
            string[] columnData = new string[] { "Col1_Row1", "Col1_Row2", "Col1_Row3", "Col1_Row4" };
            // isVertical = true => import vertically (by column)
            cells.ImportArray(columnData, firstRow: 0, firstColumn: 3, isVertical: true);

            // Save the workbook to an XLSX file
            workbook.Save("RowColumnData.xlsx");
        }
    }
}