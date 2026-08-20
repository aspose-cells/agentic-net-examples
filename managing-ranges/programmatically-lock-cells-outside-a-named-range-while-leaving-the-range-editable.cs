// Title: Lock all cells in an Aspose.Cells worksheet except a named range (C#)
// Description: C# example that creates a workbook, defines an AllowEditRange (e.g., A1:B2), locks every cell, configures worksheet protection to prevent selection of locked cells, applies a password, and saves the file.
// Keywords: Aspose.Cells lock cells | AllowEditRange C# | worksheet protection Aspose.Cells | lock entire sheet except range | C# Aspose.Cells example | protect worksheet with password | editable range Aspose.Cells | Excel cell locking C#
// Common Searches: Aspose.Cells lock cells except range | protect worksheet but allow editing of specific cells C# | AllowEditRanges example Aspose.Cells | C# lock all cells and keep A1:B2 editable | Aspose.Cells worksheet protection tutorial
// Developer Intent: Lock every cell in a worksheet while keeping a defined named range editable.
// Use Cases: Distribute a template where only the data‑entry area (e.g., A1:B2) can be edited by end users. | Generate a report that safeguards calculated cells but permits users to fill input cells within a protected range. | Create a password‑protected sheet that still allows a password‑protected editable region for specific collaborators.
// AI Prompts: Show C# code that locks all cells in an Aspose.Cells worksheet except a named editable range and protects the sheet with a password. | Provide an Aspose.Cells example using AllowEditRanges to define an editable area, then apply a locked style to the rest of the sheet before protection. | Explain how to configure worksheet protection so locked cells cannot be selected while unlocked cells remain selectable in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsLockOutsideRange
{
    // C# example that creates a workbook, defines an AllowEditRange (e.g., A1:B2), locks every cell, configures worksheet protection to prevent selection of locked cells, applies a password, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 1. Define the range that should stay editable (e.g., A1:B2)
                // ------------------------------------------------------------
                // Add an AllowEditRange for the area.
                // Parameters: name, startRow, startColumn, endRow, endColumn (zero‑based indexes)
                int editRangeIndex = sheet.AllowEditRanges.Add("EditableArea", 0, 0, 1, 1);
                // Optional: set a password for the editable range
                // sheet.AllowEditRanges[editRangeIndex].Password = "rangePwd";

                // ------------------------------------------------------------
                // 2. Lock all cells in the worksheet
                // ------------------------------------------------------------
                // Create a style with IsLocked = true
                Style lockedStyle = workbook.CreateStyle();
                lockedStyle.IsLocked = true;

                // Create a StyleFlag that applies the Locked attribute
                StyleFlag flag = new StyleFlag();
                flag.Locked = true;

                // Apply the locked style to the entire sheet
                cells.ApplyStyle(lockedStyle, flag);

                // ------------------------------------------------------------
                // 3. Protect the worksheet (all protection types)
                // ------------------------------------------------------------
                sheet.Protection.AllowSelectingLockedCell = false;
                sheet.Protection.AllowSelectingUnlockedCell = true;

                // Protect with password (optional) and all protection types.
                // The Protect method requires the old password parameter; pass an empty string if not needed.
                sheet.Protect(ProtectionType.All, "sheetPwd", string.Empty);

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "LockedOutsideEditableRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
