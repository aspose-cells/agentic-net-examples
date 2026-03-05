using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Create an Excel table from the range A1:C4
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // zero‑based index for column A
            int totalRows = 4;  // header + 3 data rows
            int totalCols = 3;  // columns A‑C

            // Add the table to the worksheet (returns the index of the new table)
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol, totalRows, totalCols, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "StudentsTable";
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;
            table.ShowTableStyleRowStripes = true;
            table.ShowTableStyleColumnStripes = false;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Insert a new row into the table (adds a row after the last data row)
            int insertRowIndex = sheet.Cells.MaxDataRow + 1; // index after current last row
            sheet.Cells.InsertRows(insertRowIndex, 1);
            sheet.Cells[insertRowIndex, 0].PutValue(4);
            sheet.Cells[insertRowIndex, 1].PutValue("Diana");
            sheet.Cells[insertRowIndex, 2].PutValue(88);

            // Update the table range to include the new row
            table.Resize(0, 0, sheet.Cells.MaxDataRow + 1, totalCols, true);

            // Delete the second data row (Bob)
            int rowToDelete = 2; // zero‑based index for row 3 (A3:C3)
            sheet.Cells.DeleteRows(rowToDelete, 1);

            // Resize the table after deletion
            table.Resize(0, 0, sheet.Cells.MaxDataRow + 1, totalCols, true);

            // Save the workbook
            string outputPath = "ManagedTableDemo.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}