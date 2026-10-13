// Title: How to password‑protect an Excel worksheet while allowing specific cells to remain editable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that protects a worksheet with a password and keeps cell B1 unlocked for user edits. | Demonstrate creating a style that sets IsLocked = false for selected cells before calling Worksheet.Protect in Aspose.Cells. | Show the steps to lock all cells by default, unlock chosen cells, apply ProtectionType.All with a custom password, and save the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# protect worksheet with password but allow editing of specific cells | unlock cells before worksheet protection Aspose.Cells .NET example | set IsLocked false style for a cell in Aspose.Cells C# | protect Excel sheet with ProtectionType.All and custom password using Aspose.Cells | save password‑protected workbook as .xlsx with Aspose.Cells C#
// Tags: worksheet protection password Aspose.Cells C# | unlock specific cells Aspose.Cells style | set IsLocked false Aspose.Cells | ProtectionType.All usage Aspose.Cells | save protected workbook xlsx Aspose.Cells

using Aspose.Cells;
using System;

// Creates a new workbook, unlocks cell B1 by applying a style with IsLocked = false, protects the worksheet with a password using ProtectionType.All, and saves the result as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add data to cells
            sheet.Cells["A1"].PutValue("This cell will be locked");
            sheet.Cells["B1"].PutValue("This cell will be editable");

            // Unlock cell B1 so it can be edited after protection
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;               // Mark as unlocked
            sheet.Cells["B1"].SetStyle(unlockedStyle);    // Apply the style

            // Protect the worksheet. Provide the new password and an empty old password.
            sheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
