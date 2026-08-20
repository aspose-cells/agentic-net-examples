// Title: C# – Remove a column from an Aspose.Cells ListObject (Excel table) without losing other data
// Description: Shows how to create a workbook, define a ListObject over range A1:C4, delete the unwanted "Name" column (column B) using Worksheet.Cells.DeleteColumn, and save the file while preserving the remaining columns.
// Keywords: Aspose.Cells delete column | C# ListObject remove column | Excel table column removal Aspose | Worksheet.Cells.DeleteColumn example | Aspose.Cells preserve data | remove column from Excel table C# | Aspose.Cells table manipulation | C# Excel column delete Aspose
// Common Searches: how to delete a column from an Aspose.Cells table in C# | remove specific column from ListObject using Aspose.Cells | Aspose.Cells delete column without affecting table data | C# code to drop a column from Excel table with Aspose | Aspose.Cells delete column B example
// Developer Intent: Delete a specific column from an Excel ListObject while keeping the other columns intact.
// Use Cases: Redact personal‑information columns before distributing a report | Trim placeholder columns after dynamic data import | Reformat generated tables by removing unnecessary fields | Prepare data sets for downstream processing by eliminating irrelevant columns
// AI Prompts: Write C# code using Aspose.Cells to delete column C from a ListObject without breaking the table. | Show how to remove multiple columns from an Aspose.Cells table while preserving all other data. | Compare Worksheet.Cells.DeleteColumn and ListObject.DeleteColumn methods in Aspose.Cells with code examples.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to create a workbook, define a ListObject over range A1:C4, delete the unwanted "Name" column (column B) using Worksheet.Cells.DeleteColumn, and save the file while preserving the remaining columns.
class RemoveColumnFromTable
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a table (columns A, B, C)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["C1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Alice");
        worksheet.Cells["C2"].PutValue(30);
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Bob");
        worksheet.Cells["C3"].PutValue(25);
        worksheet.Cells["A4"].PutValue(3);
        worksheet.Cells["B4"].PutValue("Charlie");
        worksheet.Cells["C4"].PutValue(28);

        // Create a ListObject (table) that includes the data range A1:C4
        int tableIndex = worksheet.ListObjects.Add("A1", "C4", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Delete the unwanted column (e.g., the "Name" column which is column B, index 1)
        // This uses the DeleteColumn(int) method as defined in the rule set.
        worksheet.Cells.DeleteColumn(1);

        // Save the modified workbook
        workbook.Save("TableColumnRemoved.xlsx");
    }
}
