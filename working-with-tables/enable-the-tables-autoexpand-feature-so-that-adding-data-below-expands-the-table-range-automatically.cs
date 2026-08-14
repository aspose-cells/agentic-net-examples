// Title: Auto‑Expand an Aspose.Cells ListObject When Adding New Rows – C# Example
// Description: Demonstrates how to create a worksheet table (ListObject) covering A1:B5, append rows below it, and programmatically enlarge the table range with ListObject.Resize so the table automatically includes the new data before saving the workbook.
// Keywords: Aspose.Cells auto expand table | ListObject Resize C# | expand Aspose.Cells table range | add rows to Aspose.Cells ListObject | dynamic table size .NET | C# Aspose.Cells table resizing
// Common Searches: Aspose.Cells auto expand table after adding rows | ListObject.Resize example C# | how to grow a table range in Aspose.Cells | extend Aspose.Cells ListObject programmatically | auto‑expand table Aspose.Cells .NET
// Developer Intent: Programmatically extend a ListObject’s range to include rows added beneath the original table.
// Use Cases: Generate a report where data rows are appended at runtime and the table must grow to keep formulas and formatting applied. | Create a dynamic worksheet that feeds charts or pivots, requiring the table boundaries to reflect newly inserted rows. | Automate data import processes that add batches of rows and need the table range updated before further processing.
// AI Prompts: Write C# code using Aspose.Cells to create a ListObject and automatically expand it after inserting additional rows. | Show how to calculate the new end row index and call ListObject.Resize to update the table range. | Explain how to achieve auto‑expand behavior for an Aspose.Cells table without manual resizing each time.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a worksheet table (ListObject) covering A1:B5, append rows below it, and programmatically enlarge the table range with ListObject.Resize so the table automatically includes the new data before saving the workbook.
class AutoExpandTableDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate initial data for the table (A1:B5)
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        for (int i = 2; i <= 5; i++)
        {
            cells[$"A{i}"].PutValue(i - 1);
            cells[$"B{i}"].PutValue("Item " + (i - 1));
        }

        // Add a ListObject (table) covering the range A1:B5
        int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "MyTable";

        // Add new rows below the existing table (A6:B8)
        for (int i = 6; i <= 8; i++)
        {
            cells[$"A{i}"].PutValue(i - 1);
            cells[$"B{i}"].PutValue("NewItem " + (i - 1));
        }

        // Expand the table to include the newly added rows
        // EndRow is zero‑based, so row 8 corresponds to index 7
        int newEndRow = 7;
        table.Resize(table.StartRow, table.StartColumn, newEndRow, table.EndColumn, true);

        // Save the workbook
        workbook.Save("AutoExpandTable.xlsx");
    }
}
