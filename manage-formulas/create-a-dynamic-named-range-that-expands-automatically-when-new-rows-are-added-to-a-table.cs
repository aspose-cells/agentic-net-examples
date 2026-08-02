// Title: Create a Dynamic Named Range that Auto‑Expands with a ListObject in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a workbook, add a ListObject (Excel table), define a named range using the structured reference =MyTable[#All], insert additional rows, resize the table, and retrieve the updated named range address—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | dynamic named range | structured reference | MyTable[#All] | ListObject resize | C# Excel automation | auto expanding range .NET | Excel table programmatic update | named range address | chart data source range
// Common Searches: Aspose.Cells create dynamic named range | auto expanding named range C# | resize Excel table after adding rows Aspose.Cells | structured reference MyTable[#All] usage | get updated named range address programmatically
// Developer Intent: Generate a named range that automatically grows when rows are added to an Excel table using Aspose.Cells.
// Use Cases: Maintain a continuously updating data range for formulas that must always cover the full table. | Supply a chart or pivot table with a data source that expands as new records are inserted. | Programmatically add rows to a ListObject while keeping dependent named ranges, validations, or formulas in sync.
// AI Prompts: Write C# code with Aspose.Cells that creates a ListObject, defines a named range using =MyTable[#All], adds rows, and confirms the range expands. | Show how to resize an Aspose.Cells ListObject after inserting data and retrieve the new named range address. | Explain the mechanics of the structured reference MyTable[#All] in Aspose.Cells and how it can be used for dynamic charts or formulas.

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// Demonstrates how to build a workbook, add a ListObject (Excel table), define a named range using the structured reference =MyTable[#All], insert additional rows, resize the table, and retrieve the updated named range address—all with Aspose.Cells for .NET.
class DynamicNamedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate initial data for the table (header + two rows).
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue(1);
            cells["B2"].PutValue(100);
            cells["A3"].PutValue(2);
            cells["B3"].PutValue(200);

            // Add a ListObject (Excel Table) that covers the range A1:B3.
            int tableIndex = ws.ListObjects.Add("A1", "B3", true);
            ListObject table = ws.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Create a named range that refers to the whole table using a structured reference.
            // The structured reference (=MyTable[#All]) expands automatically when rows are added.
            int nameIndex = wb.Worksheets.Names.Add("MyDynamicRange");
            Name dynName = wb.Worksheets.Names[nameIndex];
            dynName.RefersTo = "=MyTable[#All]";

            // Show the address of the named range before adding new rows.
            Aspose.Cells.Range initialRange = dynName.GetRange();
            Console.WriteLine("Initial named range address: " + initialRange.Address);

            // Add a new row to the table (after the current data rows).
            int newRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount; // first data row + existing rows
            cells[newRowIndex, 0].PutValue(3); // ID
            cells[newRowIndex, 1].PutValue(300); // Value

            // Resize the table to include the newly added row.
            // Compute the full table range (header + data) for resizing.
            int firstRow = table.DataRange.FirstRow - 1;               // header row index
            int firstColumn = table.DataRange.FirstColumn;            // first column index
            int totalRows = table.DataRange.RowCount + 1;             // data rows + header row
            int totalColumns = table.DataRange.ColumnCount;           // column count stays the same
            table.Resize(firstRow, firstColumn, totalRows, totalColumns, true);

            // Retrieve the updated range via the named range (it now includes the new row).
            Aspose.Cells.Range updatedRange = dynName.GetRange();
            Console.WriteLine("Updated named range address after adding row: " + updatedRange.Address);

            // Save the workbook.
            wb.Save("DynamicNamedRangeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
