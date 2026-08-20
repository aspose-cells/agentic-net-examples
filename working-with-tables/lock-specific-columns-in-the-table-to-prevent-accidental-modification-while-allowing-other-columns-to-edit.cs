// Title: C# – Lock Specific Columns in an Aspose.Cells Table While Allowing Others to Edit
// Description: Creates a workbook, adds a ListObject (Excel table) over A1:C3, unlocks every column, then locks the chosen columns (e.g., ID and Score) by setting the IsLocked style flag, protects the worksheet, and saves the file as LockedColumnsDemo.xlsx.
// Keywords: Aspose.Cells column lock C# | protect worksheet specific columns | unlock all columns Aspose.Cells | Excel table column protection | ListObject column lock
// Common Searches: lock only certain columns Aspose.Cells C# | protect worksheet but keep some columns editable Aspose.Cells | how to lock columns in an Excel table using Aspose.Cells
// Developer Intent: Prevent edits to selected columns in a worksheet/table while keeping the remaining columns editable.
// Use Cases: Restrict changes to ID and Score columns in a data‑entry form, allowing only the Name column to be edited. | Create a read‑only report where calculation columns are locked and input columns stay editable. | Secure financial model formulas while permitting users to modify input cells.
// AI Prompts: Generate C# code with Aspose.Cells that locks columns B and D in an existing worksheet and leaves other columns unlocked. | Show how to apply column‑level protection to a ListObject and then protect the worksheet using Aspose.Cells. | Explain how to toggle column lock states at runtime based on user selections with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a ListObject (Excel table) over A1:C3, unlocks every column, then locks the chosen columns (e.g., ID and Score) by setting the IsLocked style flag, protects the worksheet, and saves the file as LockedColumnsDemo.xlsx.
class LockSpecificColumns
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Fill some sample data that will become a table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["C2"].PutValue(85);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["C3"].PutValue(92);

        // Add a ListObject (Excel table) covering the data range
        int tableIdx = sheet.ListObjects.Add("A1", "C3", true);
        ListObject table = sheet.ListObjects[tableIdx];
        table.DisplayName = "DataTable";

        // First unlock all columns so they are editable after protection
        Style style;
        StyleFlag flag;
        for (int i = 0; i <= sheet.Cells.MaxColumn; i++)
        {
            style = sheet.Cells.Columns[i].GetStyle();
            style.IsLocked = false;          // make column unlocked
            flag = new StyleFlag();
            flag.Locked = true;               // apply the Locked flag
            sheet.Cells.Columns[i].ApplyStyle(style, flag);
        }

        // Define which columns should stay locked (e.g., column 0 and column 2)
        int[] lockedColumns = new int[] { 0, 2 };
        foreach (int col in lockedColumns)
        {
            style = sheet.Cells.Columns[col].GetStyle();
            style.IsLocked = true;           // lock this column
            flag = new StyleFlag();
            flag.Locked = true;               // apply the Locked flag
            sheet.Cells.Columns[col].ApplyStyle(style, flag);
        }

        // Protect the worksheet; locking takes effect only when protected
        sheet.Protect(ProtectionType.All);

        // Save the workbook
        wb.Save("LockedColumnsDemo.xlsx", SaveFormat.Xlsx);
    }
}
