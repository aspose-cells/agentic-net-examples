// Title: Aspose.Cells C# – Delete a ListObject column without retaining its formula
// Description: Shows how to build a workbook, define a table, set a column formula, clear the ListColumn formula, delete the column, and save the file so the removed column's calculations are not present in any remaining rows.
// Keywords: Aspose.Cells | C# | .NET | ListObject delete column | clear ListColumn formula | Excel table column removal | prevent formula copy | ListColumn.Formula | Excel automation | remove calculated column
// Common Searches: Aspose.Cells delete table column formula C# | how to remove ListObject column without leaving formulas | clear ListColumn formula before column deletion Aspose.Cells | stop formula propagation when deleting Excel column using Aspose | remove calculated column from Excel table with Aspose.Cells
// Developer Intent: Delete a ListObject column and ensure its formula does not affect other rows.
// Use Cases: Eliminate a temporary "Total" column from a sales table while keeping all other data intact. | Clean up a worksheet by removing a helper column that contained a formula, avoiding stray calculations. | Update an Excel template by discarding an obsolete column and guaranteeing no residual formulas remain.
// AI Prompts: Provide C# code that clears a ListColumn formula before deleting the column with Aspose.Cells. | Show an example of removing a ListObject column in Aspose.Cells while ensuring the column's formula is not copied to remaining rows. | Explain why setting ListColumn.Formula to an empty string stops formula propagation after the column is deleted.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to build a workbook, define a table, set a column formula, clear the ListColumn formula, delete the column, and save the file so the removed column's calculations are not present in any remaining rows.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (header + two data rows)
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Total");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(2);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(3);

        // Add a formula to the Total column (C) that references the Quantity column (B)
        sheet.Cells["C2"].Formula = "=B2*10";
        sheet.Cells["C3"].Formula = "=B3*10";

        // Create a ListObject (table) that includes the range A1:C3
        int tableIndex = sheet.ListObjects.Add("A1", "C3", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Set the same formula for the entire Total column via ListColumn.Formula
        // ListColumns are zero‑based; column C is index 2
        ListColumn totalColumn = table.ListColumns[2];
        totalColumn.Formula = "=B2*10";

        // Clear the column formula so it will not be propagated after deletion
        totalColumn.Formula = string.Empty;

        // Delete the Total column (index 2) and update references in other cells
        sheet.Cells.DeleteColumn(2, true);

        // Save the modified workbook
        workbook.Save("RemoveColumnFromListObject.xlsx");
    }
}
