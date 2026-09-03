// Title: How to protect an Excel worksheet with a password while allowing edits only in a specific unlocked range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that unlocks a defined cell range (e.g., A1:B5) and then protects the entire worksheet with a password using Aspose.Cells. | Demonstrate how to apply a style to make cells editable before calling Worksheet.Protect with ProtectionType.All in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# protect worksheet password but keep A1:B5 editable | unlock cells before worksheet protection Aspose.Cells .NET example | C# code to lock all cells except a range and save as .xlsx using Aspose.Cells | how to use Worksheet.Protect with ProtectionType.All in Aspose.Cells | set IsLocked false for a range then protect sheet Aspose.Cells C#
// Tags: worksheet protection with password Aspose.Cells | unlock cell range Aspose.Cells C# | apply unlocked style Aspose.Cells | ProtectionType.All worksheet Aspose.Cells | lock cells except specific range Aspose.Cells

using System;
using Aspose.Cells;

// // Creates a new workbook, unlocks cells A1:B5, protects the first worksheet with a password using ProtectionType.All, and saves the file as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unlock the cells that users are allowed to edit (example: range A1:B5)
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;               // cells are unlocked

            StyleFlag styleFlag = new StyleFlag();
            styleFlag.Locked = true;                       // apply the locked flag only

            sheet.Cells.CreateRange("A1:B5").ApplyStyle(unlockedStyle, styleFlag);

            // Protect the worksheet with a password; all other cells remain locked
            // The third parameter is the old password (not required here), pass null or empty string
            sheet.Protect(ProtectionType.All, "MySecurePassword", null);

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
