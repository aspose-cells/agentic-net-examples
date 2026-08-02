// Title: C# – Insert a New Row into an Aspose.Cells ListObject Table Using a Dictionary
// Description: Demonstrates how to create a workbook, define a ListObject table with headers, insert a new data row, and populate the cells by matching dictionary keys to column names with ListObject.PutCellValue, then save the file as an Excel workbook.
// Keywords: Aspose.Cells C# | ListObject InsertRow | PutCellValue example | populate Excel table from Dictionary | .NET Excel automation | add row to Aspose.Cells table | dynamic Excel data insertion | Excel table programming | Aspose.Cells tutorial USA | Aspose.Cells tutorial India
// Common Searches: Aspose.Cells add row to ListObject from dictionary | C# insert row into Excel table using Aspose.Cells | populate Aspose.Cells table with key‑value pairs | how to use ListObject.PutCellValue in C# | insert row without breaking table range Aspose.Cells
// Developer Intent: Add a new data row to an existing ListObject table and fill its cells using values from a Dictionary<string, object>.
// Use Cases: Extend an inventory sheet with products received from an API. | Append financial transaction records to a report generated at runtime. | Update a sales summary table with user‑entered items in a desktop application.
// AI Prompts: Generate C# code that inserts a row into an Aspose.Cells ListObject and assigns values from a Dictionary. | Explain how to locate a column in a ListObject by its header and set the cell value with PutCellValue. | Show the steps to insert a row in an Aspose.Cells worksheet while preserving the table range.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableDemo
{
    // Demonstrates how to create a workbook, define a ListObject table with headers, insert a new data row, and populate the cells by matching dictionary keys to column names with ListObject.PutCellValue, then save the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ----- Create a sample table (ListObject) -----
            // Header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Price");

            // Sample data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Laptop");
            cells["C2"].PutValue(999.99);

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Monitor");
            cells["C3"].PutValue(249.99);

            // Add the ListObject (table) covering A1:C3, with header row
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 2, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // ----- Dictionary containing values for the new row -----
            var newRowData = new Dictionary<string, object>
            {
                { "ID", 3 },
                { "Product", "Keyboard" },
                { "Price", 49.99 }
            };

            // ----- Insert a new row just after the existing data rows -----
            // The DataRange includes the header row, so the insertion point is:
            // first row of the table + total rows in the DataRange
            int insertRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;
            cells.InsertRow(insertRowIndex); // rule: Cells.InsertRow

            // ----- Populate the newly inserted row using the dictionary -----
            // After insertion, the new data row is the last row of the table's data range
            // Row offset within the table is (current row count - 1) because offsets are zero‑based
            int newRowOffset = table.DataRange.RowCount - 1;

            foreach (var kvp in newRowData)
            {
                // Find the column offset that matches the dictionary key (column name)
                int columnOffset = -1;
                for (int i = 0; i < table.ListColumns.Count; i++)
                {
                    if (string.Equals(table.ListColumns[i].Name, kvp.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        columnOffset = i;
                        break;
                    }
                }

                // If the column exists, put the value into the cell
                if (columnOffset >= 0)
                {
                    table.PutCellValue(newRowOffset, columnOffset, kvp.Value); // rule: ListObject.PutCellValue
                }
            }

            // Save the workbook
            workbook.Save("TableWithNewRow.xlsx");
        }
    }
}
