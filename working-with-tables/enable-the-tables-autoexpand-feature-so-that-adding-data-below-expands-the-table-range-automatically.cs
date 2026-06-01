using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableAutoExpandDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ---------- Create initial table ----------
            // Add header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");

            // Add a few data rows
            for (int i = 0; i < 3; i++)
            {
                cells[i + 1, 0].PutValue(i + 1);               // ID
                cells[i + 1, 1].PutValue($"Item {i + 1}");    // Name
            }

            // Create a ListObject (table) covering the header and the three data rows
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // ---------- Add new data below the existing table ----------
            // New rows start at row index 4 (zero‑based) because rows 0‑3 are already used
            int newRowsStart = 4;
            int rowsToAdd = 2; // number of new data rows

            for (int i = 0; i < rowsToAdd; i++)
            {
                int rowIndex = newRowsStart + i;
                cells[rowIndex, 0].PutValue(rowIndex);                 // ID
                cells[rowIndex, 1].PutValue($"NewItem {rowIndex}");    // Name
            }

            // ---------- Enable auto‑expand by resizing the table ----------
            // New end row is the last row that now contains data
            int newEndRow = newRowsStart + rowsToAdd - 1;
            // Resize the table to include the newly added rows
            // ListObject.Resize(startRow, startColumn, endRow, endColumn, hasHeaders)
            table.Resize(0, 0, newEndRow, 1, true);

            // Optional: Auto‑fit columns for better visibility
            worksheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("TableAutoExpandDemo.xlsx");
        }
    }
}