// Title: C# – Create and Auto‑Expand a ListObject from a Dynamic Range with Aspose.Cells
// Description: Shows how to detect the current data bounds, add a ListObject (Excel table) over that range, insert additional rows, and automatically resize the table using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | ListObject | dynamic range | auto resize table | Excel table expansion | MaxDataRow | MaxDataColumn | Resize method | add rows to table | Excel automation .NET
// Common Searches: Aspose.Cells create table from range | Resize ListObject after adding rows Aspose.Cells | auto expand Excel table .NET | detect last data row Aspose.Cells | C# add rows to ListObject | dynamic table range Aspose.Cells
// Developer Intent: Create a ListObject based on the existing data range and have it grow automatically when new rows are inserted.
// Use Cases: Generate an Excel table from data whose size changes at runtime without recreating the table. | Append new rows to a worksheet and keep the ListObject range up‑to‑date, preserving formatting and table name. | Export dynamically sized datasets while maintaining a named table for downstream analysis or reporting.
// AI Prompts: Provide C# code that uses Aspose.Cells to create a ListObject from a dynamic range and automatically expands it after new rows are added. | Show how to detect the last data row and column, add a table, insert additional rows, and call Resize to include them in Aspose.Cells. | Explain how to enable automatic table expansion in Aspose.Cells when data is appended later in the worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to detect the current data bounds, add a ListObject (Excel table) over that range, insert additional rows, and automatically resize the table using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate initial data (including headers)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            worksheet.Cells[i - 1, 0].PutValue(i - 1);          // ID column
            worksheet.Cells[i - 1, 1].PutValue(i * 10);       // Value column
        }

        // Determine the dynamic range that currently contains data
        int startRow = 0;                                   // first row (zero‑based)
        int startColumn = 0;                                // first column (zero‑based)
        int endRow = worksheet.Cells.MaxDataRow;            // last row with data
        int endColumn = worksheet.Cells.MaxDataColumn;      // last column with data

        // Add a ListObject (table) using the dynamic range
        int tableIndex = worksheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "DynamicTable";

        // ----- Simulate adding new rows later -----
        // Add a new data row just below the current table range
        int newRow = endRow + 1;
        worksheet.Cells[newRow, 0].PutValue(newRow);        // New ID
        worksheet.Cells[newRow, 1].PutValue(newRow * 10);  // New Value

        // Expand (resize) the ListObject to include the newly added row
        table.Resize(startRow, startColumn, newRow, endColumn, true);

        // Save the workbook (using the provided save rule)
        workbook.Save("DynamicListObject.xlsx", SaveFormat.Xlsx);
    }
}
