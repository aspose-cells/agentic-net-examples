// Title: Aspose.Cells .NET: Protect Worksheet, Allow Formatting, Block Row Insertion
// Description: Demonstrates how to enable cell, column, and row formatting while preventing row insertion on an Aspose.Cells worksheet, optionally securing it with a password and saving the workbook.
// Keywords: Aspose.Cells worksheet protection | allow formatting Aspose.Cells | disable row insertion | worksheet password .NET | custom protection settings
// Common Searches: Aspose.Cells allow formatting but prevent row insertion | protect Excel sheet with password using Aspose.Cells .NET | custom worksheet protection options Aspose.Cells | how to block row insertion in Aspose.Cells workbook
// Developer Intent: Apply worksheet protection that permits formatting actions yet blocks the addition of new rows.
// Use Cases: Distribute a template where users can style cells but cannot alter the row structure. | Share a financial report that lets recipients improve readability without changing row counts for audit integrity. | Secure a collaborative workbook with a password while allowing formatting edits and preventing accidental row inserts.
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet, enable cell/column/row formatting, disable row insertion, and set a password. | Explain how to extend the protection settings to also block column insertion while keeping formatting enabled. | Provide a step‑by‑step tutorial for applying custom worksheet protection in Aspose.Cells and saving the protected file.

using System;
using Aspose.Cells;

// Demonstrates how to enable cell, column, and row formatting while preventing row insertion on an Aspose.Cells worksheet, optionally securing it with a password and saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = sheet.Protection;

        // Allow formatting of cells, columns, and rows
        protection.AllowFormattingCell = true;
        protection.AllowFormattingColumn = true;
        protection.AllowFormattingRow = true;

        // Explicitly disallow inserting rows (default is false, set for clarity)
        protection.AllowInsertingRow = false;

        // Optional: set a password for the protection
        protection.Password = "MySecretPassword";

        // Apply protection to the worksheet (protect all aspects)
        sheet.Protect(ProtectionType.All);

        // Save the workbook with the applied protection
        workbook.Save("AdvancedWorksheetProtection.xlsx");
    }
}
