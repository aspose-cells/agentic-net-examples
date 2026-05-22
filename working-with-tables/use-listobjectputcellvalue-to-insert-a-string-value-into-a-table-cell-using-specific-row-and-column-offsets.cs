using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectPutCellValueDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header row for the table
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Description");

            // Add a few data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue("Initial description");

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue("Initial description");

            // Create a ListObject (table) that covers the data range including the header
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 2, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Use PutCellValue to insert a string into a specific cell of the table
            // rowOffset = 1 (second data row, because 0 is the header row)
            // columnOffset = 2 (third column, "Description")
            // value = new string to set
            table.PutCellValue(rowOffset: 1, columnOffset: 2, value: "Updated description for Bob");

            // Save the workbook to a file
            workbook.Save("ListObjectPutCellValueDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}