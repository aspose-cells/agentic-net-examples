using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the leftmost two columns (A and B)
            // Row index = 0 (no frozen rows), Column index = 2 (freeze before column C)
            // freezedRows = 0, freezedColumns = 2
            worksheet.FreezePanes(0, 2, 0, 2);

            // Save the workbook
            workbook.Save("FreezeLeftTwoColumns.xlsx");
        }
    }
}