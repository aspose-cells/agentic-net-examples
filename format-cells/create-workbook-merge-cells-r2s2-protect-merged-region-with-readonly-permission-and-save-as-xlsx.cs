// Title: Create an XLSX workbook, merge cells R2:S2, lock the merged range, and protect the worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells merge cells R2:S2 and set read‑only | how to lock only a merged range in an Excel sheet using Aspose.Cells C# | save a worksheet with a protected merged area to XLSX with Aspose.Cells | unlock all cells then lock a specific merged region in Aspose.Cells .NET example | apply password‑based worksheet protection for locked cells using Aspose.Cells
// Tags: merged range lock Aspose.Cells | global cell unlock selective lock C# | export locked worksheet to XLSX Aspose | password protected worksheet .NET | cell style lock flag Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, unlocks every cell, merges the range R2:S2, locks that merged region, protects the worksheet with a password, and saves the file as MergedProtected.xlsx in XLSX format using Aspose.Cells for .NET.
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

            // Unlock all cells so that only the merged region will be read‑only
            Style unlockedStyle = sheet.Cells.Style;
            unlockedStyle.IsLocked = false;
            // Apply the unlocked style to all cells (Lock flag indicates which attribute to apply)
            sheet.Cells.ApplyStyle(unlockedStyle, new StyleFlag { Locked = true });

            // Merge cells R2:S2 (row index 1, column indices 17 to 18)
            sheet.Cells.Merge(1, 17, 1, 18);

            // Lock the merged region
            Style lockedStyle = sheet.Cells[1, 17].GetStyle();
            lockedStyle.IsLocked = true;
            sheet.Cells[1, 17].SetStyle(lockedStyle);

            // Protect the worksheet (read‑only for locked cells)
            sheet.Protect(ProtectionType.All, "pwd123", string.Empty);

            // Save the workbook as XLSX
            workbook.Save("MergedProtected.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
