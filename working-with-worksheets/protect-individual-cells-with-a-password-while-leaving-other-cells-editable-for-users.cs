// Title: How to password‑protect an Excel worksheet while keeping selected cells editable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that locks cell A1, unlocks cell B1, and protects the worksheet with a password. | Generate an Aspose.Cells .NET snippet that sets IsLocked = false for chosen cells before applying worksheet protection. | Create a C# example demonstrating how to unlock particular cells, then call Worksheet.Protect to enforce read‑only on locked cells.
// Common Searches: Aspose.Cells C# lock only certain cells and protect sheet with password | How to unlock specific cells in an Excel file while protecting the rest using Aspose.Cells .NET | C# example for worksheet protection with password and selective cell editing in Aspose.Cells | Set cell IsLocked property in Aspose.Cells before applying worksheet protection
// Tags: worksheet protection password Aspose.Cells .NET | unlock specific cells Aspose.Cells | cell IsLocked style C# | selective cell edit Aspose.Cells | protect Excel sheet Aspose.Cells example

using System;
using Aspose.Cells;

// The program creates a new workbook, writes values to A1 (locked) and B1 (editable), sets B1's style IsLocked to false, protects the worksheet with a password using Worksheet.Protect, and saves the file as ProtectedCells.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example data: A1 will be locked, B1 will be editable
            sheet.Cells["A1"].PutValue("Locked Cell");
            sheet.Cells["B1"].PutValue("Editable Cell");

            // By default all cells are locked. Unlock B1 so users can edit it.
            Style unlockedStyle = sheet.Cells["B1"].GetStyle();
            unlockedStyle.IsLocked = false;               // make the cell editable
            sheet.Cells["B1"].SetStyle(unlockedStyle);    // apply the style to the cell

            // Protect the worksheet with a password.
            // All locked cells become read‑only, while unlocked cells (B1) remain editable.
            // The third parameter (oldPassword) is not required for new protection; pass null.
            sheet.Protect(ProtectionType.All, "MySecurePassword", null);

            // Save the workbook to a file
            workbook.Save("ProtectedCells.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
