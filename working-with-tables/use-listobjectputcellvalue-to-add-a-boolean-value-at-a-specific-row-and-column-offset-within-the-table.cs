using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectBooleanDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add header row for the table
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("IsActive"); // Boolean column

            // Add some sample data rows
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["C2"].PutValue(false); // initial value

            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");
            worksheet.Cells["C3"].PutValue(false); // initial value

            // Create a ListObject (table) that includes the data range A1:C3
            int tableIndex = worksheet.ListObjects.Add("A1", "C3", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Update the boolean value in the second data row (row offset 2, column offset 2)
            // Row offset is zero‑based relative to the first data row (excluding header)
            // Column offset is zero‑based relative to the first column of the table
            table.PutCellValue(rowOffset: 2, columnOffset: 2, value: true);

            // Save the workbook to a file
            workbook.Save("ListObjectBooleanDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}