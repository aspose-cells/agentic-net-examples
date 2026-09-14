// Title: Password‑protect specific cells while leaving other cells editable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that locks selected cells, unlocks others, and applies a worksheet password. | Show how to create a style that sets IsLocked = false and use it to make certain cells editable before protecting the sheet with a password.
// Common Searches: Aspose.Cells C# protect only certain cells with a password | How to unlock specific cells while protecting the rest of the worksheet in Aspose.Cells | C# example for cell‑level protection using Aspose.Cells workbook | Apply password protection to a worksheet but allow editing of designated cells in .NET | Selective cell locking with Aspose.Cells and password protection
// Tags: worksheet password protection Aspose.Cells | unlock cells for editing Aspose.Cells | cell lock style Aspose.Cells C# | selective cell protection .NET | apply locked property style Aspose.Cells

using System;
using Aspose.Cells;

// Creates a workbook, writes data to A1‑B2, unlocks cells B1 and B2 via a style with IsLocked = false, protects the worksheet with a password so locked cells become read‑only, and saves the file as ProtectedCells.xlsx.
class ProtectCellsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some sample data
            sheet.Cells["A1"].PutValue("Read‑Only Cell");
            sheet.Cells["B1"].PutValue("Editable Cell");
            sheet.Cells["A2"].PutValue("Read‑Only Cell 2");
            sheet.Cells["B2"].PutValue("Editable Cell 2");

            // By default all cells are locked. Unlock the cells that should remain editable.
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // make editable

            // Apply the unlocked style to the cells that users can edit
            StyleFlag flag = new StyleFlag();
            flag.Locked = true; // indicate that the Locked property should be applied

            sheet.Cells["B1"].SetStyle(unlockedStyle, flag);
            sheet.Cells["B2"].SetStyle(unlockedStyle, flag);

            // Protect the worksheet with a password.
            // All locked cells become read‑only, while unlocked cells stay editable.
            sheet.Protect(ProtectionType.All, "MySecretPassword", null);

            // Save the workbook
            workbook.Save("ProtectedCells.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
