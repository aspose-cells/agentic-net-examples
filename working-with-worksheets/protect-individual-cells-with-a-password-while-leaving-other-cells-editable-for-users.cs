// Title: How to lock specific cells with a password while keeping other cells editable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock only selected cells and set worksheet password | How to make some Excel cells editable after protecting the sheet with Aspose.Cells | C# example for unlocking a cell before applying worksheet protection in Aspose.Cells | Set IsLocked false for a cell and protect worksheet with password using Aspose.Cells | Protect worksheet but allow editing of specific cells in .NET with Aspose.Cells
// Tags: cell-level lock Aspose.Cells C# | unlock cell style IsLocked false Aspose.Cells | worksheet password protection Aspose.Cells | protect specific cells Excel C# | Aspose.Cells selective cell protection | Excel cell lock/unlock with Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new workbook, writes values to A1 (locked) and B1 (editable), applies a style with IsLocked = false to B1, then protects the worksheet with a password using Protect(ProtectionType.All, "MyPassword", null), and saves the file as ProtectedCells.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example data: A1 will be locked, B1 will be editable
            sheet.Cells["A1"].PutValue("Locked Cell");
            sheet.Cells["B1"].PutValue("Editable Cell");

            // By default all cells are locked. Create a style that unlocks cells.
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;

            // Apply the unlocked style to the cells that should remain editable.
            sheet.Cells["B1"].SetStyle(unlockedStyle);

            // Protect the worksheet with a password.
            // The overload requires an old password parameter; pass null when not needed.
            sheet.Protect(ProtectionType.All, "MyPassword", null);

            // Save the workbook to a file.
            workbook.Save("ProtectedCells.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
