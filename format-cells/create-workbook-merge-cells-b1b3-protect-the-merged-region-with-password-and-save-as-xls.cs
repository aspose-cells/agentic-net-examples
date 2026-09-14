// Title: Create an XLS workbook, merge cells B1:B3, lock only the merged range, and protect the sheet with a password using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that merges the range B1:B3, keeps all other cells unlocked, locks the merged cells, applies full worksheet protection with a password, and saves the file as an Excel 97‑2003 (.xls) document. | Show how to apply an unlocked style to every cell, then apply a locked style just to a merged region and enable password‑based protection for the worksheet in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# merge B1:B3 and lock only that range while other cells stay editable | how to protect a worksheet with a password and save as Excel 97-2003 using Aspose.Cells | C# example to unlock all cells, lock a merged area, and apply sheet protection in Aspose.Cells
// Tags: merge cells lock range Aspose.Cells C# | password protect worksheet Excel97To2003 Aspose.Cells | apply unlocked style whole sheet Aspose.Cells | lock specific merged region Aspose.Cells

using System;
using Aspose.Cells;

// // This C# program creates a new workbook, merges cells B1:B3, unlocks all cells, locks the merged region, protects the worksheet with a password, and saves the result as an Excel 97‑2003 (.xls) file using Aspose.Cells for .NET.
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

            // Merge cells B1:B3 (row index 1, column index 1, 3 rows, 1 column)
            sheet.Cells.Merge(1, 1, 3, 1);

            // Unlock all cells in the worksheet
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;
            StyleFlag unlockedFlag = new StyleFlag();
            unlockedFlag.Locked = true; // Apply the Locked property from the style
            sheet.Cells.ApplyStyle(unlockedStyle, unlockedFlag);

            // Lock the merged region (B1:B3)
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true;
            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true; // Apply the Locked property from the style
            sheet.Cells.CreateRange("B1:B3").ApplyStyle(lockedStyle, lockedFlag);

            // Protect the worksheet with a password (oldPassword is not required, pass empty string)
            sheet.Protect(ProtectionType.All, "MyPassword", string.Empty);

            // Save the workbook as XLS (Excel 97-2003 format)
            workbook.Save("MergedProtected.xls", SaveFormat.Excel97To2003);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
