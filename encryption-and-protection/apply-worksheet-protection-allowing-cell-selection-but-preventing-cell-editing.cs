// Title: Aspose.Cells .NET – Protect Worksheet While Enabling Cell Selection and Blocking Editing
// Description: Shows how to create a workbook, retrieve the first worksheet, set protection flags so users can select any cell, prevent content changes, apply full protection without a password, and save the file as WorksheetProtected.xlsx.
// Keywords: Aspose.Cells worksheet protection | C# protect Excel sheet | allow cell selection Aspose | disable editing locked cells | no‑password worksheet protection | ProtectionType.All | Aspose.Cells .NET API | read‑only workbook
// Common Searches: aspnet protect excel sheet but allow selection | aspose.cells stop editing while still letting users select cells | worksheet protection without password in C# | enable selecting locked cells Aspose.Cells | set worksheet protection options programmatically
// Developer Intent: Add worksheet‑level protection that permits cell selection but forbids any modifications to the sheet’s data.
// Use Cases: Distribute a template where only specific input cells are editable. | Publish a financial report that must stay unchanged while users can copy values. | Create a read‑only dashboard for external stakeholders using Aspose.Cells.
// AI Prompts: Generate C# code to protect an Aspose.Cells worksheet, allow selection of all cells, and require a password. | Show how to toggle individual protection flags (select locked, select unlocked, edit content) in Aspose.Cells for .NET. | Explain the difference between ProtectionType.All and selective protection types in Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // Shows how to create a workbook, retrieve the first worksheet, set protection flags so users can select any cell, prevent content changes, apply full protection without a password, and save the file as WorksheetProtected.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access the worksheet's protection settings
            Protection protection = sheet.Protection;

            // Allow users to select both locked and unlocked cells
            protection.AllowSelectingLockedCell = true;
            protection.AllowSelectingUnlockedCell = true;

            // Prevent editing of locked cells
            protection.AllowEditingContent = false;

            // Apply protection to the worksheet (no password required)
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("WorksheetProtected.xlsx");
        }
    }
}
