using System;
using Aspose.Cells;

class FreezeFirstRowColumn
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze the first row and first column simultaneously.
        // Row index = 1, Column index = 1, freeze 1 row and 1 column.
        worksheet.FreezePanes(1, 1, 1, 1);

        // Save the workbook
        workbook.Save("FreezeFirstRowColumn.xlsx");
    }
}