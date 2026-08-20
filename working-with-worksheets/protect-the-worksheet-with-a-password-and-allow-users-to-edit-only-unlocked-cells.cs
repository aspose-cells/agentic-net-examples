// Title: Protect Aspose.Cells worksheet with a password while keeping B2:C3 editable (C#)
// Description: Creates a new Workbook, unlocks the range B2:C3 by setting IsLocked = false, configures worksheet protection (disallowing edits to locked cells, allowing selection of both locked and unlocked cells), applies a password, protects the sheet with ProtectionType.All, and saves the file as ProtectedWorksheet.xlsx.
// Keywords: Aspose.Cells C# worksheet protection | password protect worksheet Aspose.Cells | unlock specific cells Aspose.Cells | IsLocked false style | ProtectionType.All | editable range B2:C3 | allow editing unlocked cells | worksheet security C#
// Common Searches: Aspose.Cells protect worksheet with password C# | unlock cells before protecting sheet Aspose.Cells | keep B2:C3 editable after worksheet protection | how to set IsLocked false in Aspose.Cells | worksheet protection options AllowSelectingLockedCell
// Developer Intent: Protect a worksheet with a password but let users modify only the unlocked cells.
// Use Cases: Distribute a template where only input cells (e.g., B2:C3) are editable while formulas stay locked. | Share a financial model that prevents accidental changes to calculations but allows data entry in designated ranges. | Create a report that blocks formatting changes yet permits users to update result cells within a specific area.
// AI Prompts: Generate C# code using Aspose.Cells to password‑protect a worksheet and keep range B2:C3 editable. | Show how to apply an unlocked style (IsLocked = false) to a cell range and then protect the sheet with ProtectionType.All. | Explain the worksheet protection settings in Aspose.Cells, such as AllowSelectingLockedCell and AllowSelectingUnlockedCell.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Creates a new Workbook, unlocks the range B2:C3 by setting IsLocked = false, configures worksheet protection (disallowing edits to locked cells, allowing selection of both locked and unlocked cells), applies a password, protects the sheet with ProtectionType.All, and saves the file as ProtectedWorksheet.xlsx.
    public class ProtectWorksheetWithUnlockedCells
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Unlock cells B2:C3 so users can edit them after protection
                // By default all cells are locked; we need to set IsLocked = false for the range we want editable
                Aspose.Cells.Range unlockRange = sheet.Cells.CreateRange("B2:C3");
                Style unlockedStyle = workbook.CreateStyle();
                unlockedStyle.IsLocked = false; // Unlock the cells

                // Apply the unlocked style to the range (apply all style attributes)
                unlockRange.ApplyStyle(unlockedStyle, new StyleFlag { All = true });

                // Configure protection options
                Protection protection = sheet.Protection;
                protection.AllowEditingContent = false;       // Disallow editing of locked cells
                protection.AllowSelectingLockedCell = true;   // Optional: allow selection of locked cells
                protection.AllowSelectingUnlockedCell = true; // Optional: allow selection of unlocked cells
                protection.Password = "MySecretPassword";     // Set the worksheet password

                // Apply protection to the worksheet (lifecycle rule: protect)
                // Using Protect with ProtectionType.All applies all protection options
                sheet.Protect(ProtectionType.All);

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ProtectedWorksheet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorksheetWithUnlockedCells.Run();
        }
    }
}
