using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectHeaderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data including initial header values
            cells["A1"].PutValue("OldHeader1");
            cells["B1"].PutValue("OldHeader2");
            cells["A2"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["B3"].PutValue(40);

            // Add a ListObject (table) that includes the header row
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Use PutCellValue with row and column offsets to change a header cell.
            // Row offset 0 = header row, column offset 0 = first column.
            table.PutCellValue(0, 0, "NewHeader1");
            table.PutCellValue(0, 1, "NewHeader2");

            // Save the workbook to a file
            workbook.Save("ListObjectHeaderUpdated.xlsx", SaveFormat.Xlsx);
        }
    }
}