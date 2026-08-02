// Title: Password‑protect an Excel worksheet and allow selection of locked cells only – Aspose.Cells C# example
// Description: Shows how to set a worksheet password, enable selection of locked cells while disabling selection of unlocked cells, apply full protection (ProtectionType.All), and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | worksheet protection | password protection | allow selecting locked cells | ProtectionType.All | Excel security | protect worksheet programmatically | locked cell selection | Aspose.Cells example
// Common Searches: Aspose.Cells protect worksheet with password C# | allow selecting only locked cells Aspose.Cells | worksheet.Protect ProtectionType.All sample | set worksheet protection options Aspose.Cells .NET | restrict cell selection in Excel using Aspose.Cells
// Developer Intent: Apply a password to a worksheet and limit user selection to locked cells only.
// Use Cases: Distribute a read‑only report where users can navigate locked cells but cannot edit any content. | Provide a template that safeguards formulas and formatting while permitting selection of locked cells for review. | Secure a workbook before publishing, ensuring only locked cells are selectable and the file remains password‑protected.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet with a password and enables selection of locked cells only. | Explain how to modify the protection settings to also allow selecting unlocked cells while keeping the worksheet password protected. | Show an example of combining multiple ProtectionType flags (e.g., Objects, Scenarios) with a password in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Shows how to set a worksheet password, enable selection of locked cells while disabling selection of unlocked cells, apply full protection (ProtectionType.All), and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the worksheet's protection settings
            Protection protection = worksheet.Protection;

            // Allow users to select locked cells only
            protection.AllowSelectingLockedCell = true;
            protection.AllowSelectingUnlockedCell = false; // optional, default is false

            // Set a password for the worksheet protection
            protection.Password = "myPassword";

            // Apply protection to the worksheet (all protection types)
            worksheet.Protect(ProtectionType.All);

            // Save the protected workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
    }
}
