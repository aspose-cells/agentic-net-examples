// Title: Get Excel Table Display Name with Aspose.Cells for .NET (Cell.GetTable)
// Description: Loads a workbook, selects a cell inside a table, calls Cell.GetTable() to obtain the ListObject, and reads its DisplayName property, handling the case where the cell is not part of any table.
// Keywords: Aspose.Cells | C# | Cell.GetTable | ListObject | DisplayName | Excel table name | retrieve table name | read table properties | Aspose.Cells .NET example | Excel table metadata
// Common Searches: Aspose.Cells get table name C# | Cell.GetTable display name example | How to read Excel table display name using Aspose | Check if a cell belongs to a table Aspose.Cells | Retrieve ListObject DisplayName Aspose.Cells
// Developer Intent: Identify the display name of the Excel table that contains a specified cell.
// Use Cases: Log the table name for a cell during data‑validation workflows. | Confirm that a target cell resides in the expected table before extracting rows. | Generate a quick inventory of tables in a worksheet by probing representative cells and outputting each DisplayName.
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, picks a cell, uses GetTable to fetch its ListObject, and prints the table's DisplayName. | Provide an Aspose.Cells .NET snippet that checks whether a given cell is inside a table and returns the table's display name or a not‑found message.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads a workbook, selects a cell inside a table, calls Cell.GetTable() to obtain the ListObject, and reads its DisplayName property, handling the case where the cell is not part of any table.
class GetTableDisplayName
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Choose a cell that is inside a table; for example, cell A2
        Cell cell = worksheet.Cells["A2"];

        // Retrieve the table (ListObject) that the cell belongs to
        ListObject table = cell.GetTable();

        if (table != null)
        {
            // Read and display the table's display name
            Console.WriteLine("Table Display Name: " + table.DisplayName);
        }
        else
        {
            Console.WriteLine("The specified cell does not belong to any table.");
        }

        // Save the workbook if any changes were made (optional)
        workbook.Save("output.xlsx");
    }
}
