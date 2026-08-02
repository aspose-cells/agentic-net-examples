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

            // Populate initial data for the table (including header row)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            for (int i = 2; i <= 5; i++) // rows 2‑5 contain data
            {
                cells[$"A{i}"].PutValue(i - 1);
                cells[$"B{i}"].PutValue($"Item {i - 1}");
            }

            // Add a ListObject (table) covering the initial range A1:B5 (header + 4 data rows)
            int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // -----------------------------------------------------------------
            // Simulate adding new data rows below the current table range.
            // The table will not automatically grow, so we must resize it.
            // -----------------------------------------------------------------
            int newRowsCount = 3; // number of rows we want to add
            int startInsertRow = table.EndRow + 1; // first row after the current table

            // Insert blank rows to make space for new data
            cells.InsertRows(startInsertRow, newRowsCount);

            // Fill the newly inserted rows with data
            for (int i = 0; i < newRowsCount; i++)
            {
                int rowIdx = startInsertRow + i;
                cells[rowIdx, 0].PutValue(table.EndRow - 3 + i + 1); // simple incremental ID
                cells[rowIdx, 1].PutValue($"NewItem {i + 1}");
            }

            // After inserting rows, resize the table to include the new data range.
            // Resize parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            table.Resize(
                table.StartRow,               // keep original start row
                table.StartColumn,            // keep original start column
                table.EndRow + newRowsCount, // new end row after added rows
                table.EndColumn,              // keep original end column
                true);                       // table still has a header row

            // Optional: Auto‑fit rows and columns for better appearance
            worksheet.AutoFitRows();
            worksheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("TableAutoExpandDemo.xlsx");
        }
    }
}