// Title: Remove a column from an Aspose.Cells ListObject (Excel table) in C# while preserving other data
// Description: C# example that creates a workbook, builds a ListObject spanning A1:C3, deletes the "Name" column (column B) using sheet.Cells.DeleteColumn, automatically shifts remaining columns left, updates table references, and saves the result.
// Keywords: Aspose.Cells delete column | C# ListObject remove column | Aspose.Cells DeleteColumn method | Excel table column removal .NET | preserve table data Aspose.Cells | shift columns left Aspose.Cells | remove unwanted column C#
// Common Searches: How to delete a column from an Aspose.Cells ListObject | C# Aspose.Cells remove specific column from Excel table | Delete column B in Aspose.Cells and keep other data | Remove unwanted column from Excel workbook using Aspose.Cells .NET | Aspose.Cells DeleteColumn shift columns left
// Developer Intent: Delete a selected column from an Excel table without losing the data in the remaining columns.
// Use Cases: Strip confidential fields (e.g., names) before sharing a report | Clean up generated spreadsheets by removing temporary calculation columns | Prepare data for import by dropping unnecessary columns from a table
// AI Prompts: Generate C# code with Aspose.Cells that removes the "Name" column from a ListObject and keeps the table structure intact. | Show how to delete a column by zero‑based index in Aspose.Cells and automatically update table references and formulas. | Provide an example of removing multiple columns from an Aspose.Cells ListObject while preserving the remaining data.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// C# example that creates a workbook, builds a ListObject spanning A1:C3, deletes the "Name" column (column B) using sheet.Cells.DeleteColumn, automatically shifts remaining columns left, updates table references, and saves the result.
class RemoveTableColumn
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for a table (ID, Name, Age)
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Age");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["C3"].PutValue(25);

        // Create a ListObject (Excel table) that spans A1:C3
        int tableIndex = sheet.ListObjects.Add("A1", "C3", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Remove the unwanted column "Name" (column B, zero‑based index 1)
        // DeleteColumn shifts remaining columns left and updates references.
        sheet.Cells.DeleteColumn(1, true);

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
