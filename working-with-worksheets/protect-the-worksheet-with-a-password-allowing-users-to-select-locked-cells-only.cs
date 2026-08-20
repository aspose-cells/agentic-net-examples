// Title: Protect an Excel worksheet with a password and allow only locked‑cell selection using Aspose.Cells for .NET
// Description: Shows how to assign a password to an Aspose.Cells worksheet, enable selection of locked cells while disabling selection of unlocked cells, apply full protection (ProtectionType.All), and save the workbook.
// Keywords: Aspose.Cells | C# | worksheet protection | password protection | allow selecting locked cells | ProtectionType.All | Excel security | read‑only worksheet
// Common Searches: Aspose.Cells protect worksheet password C# | allow selecting locked cells Aspose.Cells | worksheet protection options .NET | disable unlocked cell selection Aspose.Cells | set worksheet password Aspose.Cells
// Developer Intent: Add password protection to a worksheet while restricting selection to locked cells.
// Use Cases: Distribute a read‑only financial report where users can click locked cells to view formulas. | Provide a template that lets users navigate locked cells for guidance but prevents any edits. | Secure sensitive data in a shared workbook, allowing only cell selection for review. | Create a protected dashboard where users can select chart cells without modifying underlying data.
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet, set a password, and enable only locked‑cell selection. | Show how to configure Aspose.Cells Protection to also allow selecting unlocked cells while keeping the sheet password protected. | Provide an example that applies specific ProtectionType flags (Objects, Scenarios) together with a password in Aspose.Cells. | Generate a snippet that unprotects a worksheet, changes its protection settings, and re‑protects it with a new password.

using System;
using Aspose.Cells;

// Shows how to assign a password to an Aspose.Cells worksheet, enable selection of locked cells while disabling selection of unlocked cells, apply full protection (ProtectionType.All), and save the workbook.
class ProtectWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = sheet.Protection;

        // Allow users to select locked cells only
        protection.AllowSelectingLockedCell = true;
        protection.AllowSelectingUnlockedCell = false; // optional, default is false

        // Set a password for the worksheet
        protection.Password = "mySecretPwd";

        // Apply protection (all protection types) to the worksheet
        sheet.Protect(ProtectionType.All);

        // Save the protected workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
