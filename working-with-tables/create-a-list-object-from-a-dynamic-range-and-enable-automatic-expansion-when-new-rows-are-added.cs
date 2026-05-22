using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including headers) in a dynamic range
            // A1:C4 will be the initial range; later rows can be added below it
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(90);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Add a ListObject (Excel table) using the range that contains the data
            // The range is defined by its start and end cell addresses
            ListObjectCollection listObjects = sheet.ListObjects;
            int tableIndex = listObjects.Add("A1", "C4", true); // hasHeaders = true
            ListObject table = listObjects[tableIndex];

            // Optional: set a table style for better visual appearance
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Demonstrate automatic expansion:
            // Adding a new row directly below the current table expands it automatically
            // Use PutCellValue on the row index that is one past the current EndRow
            int newRow = table.EndRow + 1; // row index where the new data will be placed
            table.PutCellValue(newRow, 0, 4);               // ID
            table.PutCellValue(newRow, 1, "Diana");         // Name
            table.PutCellValue(newRow, 2, 92);              // Score

            // At this point the ListObject has automatically grown to include the new row

            // Save the workbook (lifecycle rule compliance)
            workbook.Save("DynamicListObject.xlsx", SaveFormat.Xlsx);
        }
    }
}