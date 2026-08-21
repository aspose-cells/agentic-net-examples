// Title: Aspose.Cells for .NET – Protect Worksheet with Password and Allow Only Locked‑Cell Selection (C#)
// Description: C# example that creates a workbook, sets a protection password (Secure123), enables selection of locked cells while disabling selection of unlocked cells, applies full protection (ProtectionType.All) and saves the file as ProtectedWorksheet.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells protect worksheet C# | Excel sheet password protection .NET | allow selecting locked cells Aspose | worksheet protection options | ProtectionType.All Aspose.Cells | secure Excel template C#
// Common Searches: how to protect an Excel worksheet with a password using Aspose.Cells | allow only locked‑cell selection in protected sheet Aspose | C# Aspose.Cells worksheet protection settings | set worksheet password and selection rules .NET | protect multiple worksheets with different passwords Aspose
// Developer Intent: Apply password protection to a worksheet while permitting users to select only the locked cells.
// Use Cases: Distribute a read‑only template where users can copy data from locked cells but cannot edit any content. | Create a report that lets viewers navigate highlighted locked cells without exposing editable ranges. | Secure a shared workbook so only locked cells are selectable for copy‑paste, preventing changes to unlocked areas.
// AI Prompts: Generate C# code with Aspose.Cells to protect a worksheet using password 'Secure123' and allow selection of locked cells only. | Show how to modify the protection object to also enable selection of unlocked cells in Aspose.Cells. | Provide an example that protects several worksheets in one workbook, each with a different password, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// C# example that creates a workbook, sets a protection password (Secure123), enables selection of locked cells while disabling selection of unlocked cells, applies full protection (ProtectionType.All) and saves the file as ProtectedWorksheet.xlsx using Aspose.Cells.
class ProtectWorksheetExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the worksheet protection settings
        Protection protection = worksheet.Protection;

        // Allow users to select locked cells only
        protection.AllowSelectingLockedCell = true;
        protection.AllowSelectingUnlockedCell = false; // optional, default is false

        // Set the protection password
        protection.Password = "Secure123";

        // Apply protection to the worksheet (all protection types)
        worksheet.Protect(ProtectionType.All);

        // Save the protected workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
