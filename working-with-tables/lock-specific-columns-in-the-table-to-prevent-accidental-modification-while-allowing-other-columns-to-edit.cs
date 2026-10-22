// Title: How to lock specific columns in an Excel worksheet while keeping other columns editable using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that first sets all cells to editable, then restricts columns A and B, protects the sheet with a password, and saves the workbook as an .xlsx file. | Show how to define a style with IsLocked=true, assign it to chosen columns, and then enable full worksheet protection while allowing edits on the remaining cells in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# lock only certain columns in a worksheet and protect the sheet with a password | How to make columns read‑only in an Excel file using Aspose.Cells while other cells stay editable | C# example for unlocking all cells then applying locked style to column A and B with worksheet protection in Aspose.Cells
// Tags: lock specific columns Aspose.Cells C# | worksheet protection with password Aspose.Cells | set IsLocked property on column range Aspose.Cells | editable cells except locked columns Aspose.Cells | read‑only columns Excel Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a workbook, fills sample data, makes every cell editable, applies a locked style to columns A and B, protects the worksheet with a password, and saves the file as LockedColumns.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for demonstration
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // -------------------------------------------------
            // Step 1: Unlock all cells (default is locked)
            // -------------------------------------------------
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // make cells editable
            StyleFlag unlockedFlag = new StyleFlag();
            unlockedFlag.Locked = true;      // apply the Locked property
            sheet.Cells.ApplyStyle(unlockedStyle, unlockedFlag);

            // -------------------------------------------------
            // Step 2: Lock specific columns (e.g., columns A and B)
            // -------------------------------------------------
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true; // make cells read‑only
            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true;    // apply the Locked property

            // Apply locked style to column A (index 0) and column B (index 1)
            sheet.Cells.Columns[0].ApplyStyle(lockedStyle, lockedFlag);
            sheet.Cells.Columns[1].ApplyStyle(lockedStyle, lockedFlag);

            // -------------------------------------------------
            // Step 3: Protect the worksheet so locked cells cannot be edited
            // -------------------------------------------------
            // Old password is empty because the sheet is not previously protected
            sheet.Protect(ProtectionType.All, string.Empty, "StrongPassword!123");

            // -------------------------------------------------
            // Step 4: Save the workbook
            // -------------------------------------------------
            workbook.Save("LockedColumns.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
