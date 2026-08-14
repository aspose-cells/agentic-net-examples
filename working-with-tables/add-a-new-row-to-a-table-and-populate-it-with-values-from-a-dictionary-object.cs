// Title: Add a Row to an Aspose.Cells ListObject Table from a C# Dictionary
// Description: Shows how to build a workbook, define a ListObject table with headers, and append a new row by matching dictionary keys to column names. Values are written with PutCellValue and the workbook is saved as TableWithNewRow.xlsx.
// Keywords: Aspose.Cells | ListObject | add row | C# dictionary | PutCellValue | Excel table population | .NET | dynamic column mapping | workbook save | Excel automation
// Common Searches: Aspose.Cells add row to ListObject using dictionary | C# populate Excel table from Dictionary<string, object> | How to insert a new record into an Aspose.Cells table | PutCellValue example with column headers | Append data to Excel ListObject programmatically
// Developer Intent: Append a new record to an existing Aspose.Cells ListObject by mapping dictionary entries to the table's column headers.
// Use Cases: Insert product details stored in a Dictionary into an inventory table. | Add a generated financial entry to a reporting table without hard‑coding column positions. | Synchronize deserialized JSON objects with an Excel table by iterating over key/value pairs.
// AI Prompts: Write C# code that adds a row to an Aspose.Cells ListObject using a Dictionary<string, object> where keys match the table headers. | Explain how to handle missing keys in the dictionary when populating a new ListObject row, leaving those cells empty. | Provide a loop that inserts multiple rows from a List<Dictionary<string, object>> into an Aspose.Cells table.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to build a workbook, define a ListObject table with headers, and append a new row by matching dictionary keys to column names. Values are written with PutCellValue and the workbook is saved as TableWithNewRow.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define table headers
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Price");

        // Add some initial data rows
        cells["A2"].PutValue(1);
        cells["B2"].PutValue("Apple");
        cells["C2"].PutValue(0.5);
        cells["A3"].PutValue(2);
        cells["B3"].PutValue("Banana");
        cells["C3"].PutValue(0.3);

        // Create a ListObject (Excel table) that includes the header and data rows
        int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Dictionary containing values for the new row
        var newRowValues = new Dictionary<string, object>
        {
            { "ID", 3 },
            { "Name", "Cherry" },
            { "Price", 0.8 }
        };

        // Determine the offset for the new row (after existing data rows)
        int newRowOffset = table.DataRange.RowCount; // zero‑based offset within the table

        // Populate the new row using the dictionary values
        for (int col = 0; col < table.ListColumns.Count; col++)
        {
            string header = table.ListColumns[col].Name;
            if (newRowValues.TryGetValue(header, out object value))
            {
                table.PutCellValue(newRowOffset, col, value);
            }
        }

        // Save the workbook
        workbook.Save("TableWithNewRow.xlsx", SaveFormat.Xlsx);
    }
}
