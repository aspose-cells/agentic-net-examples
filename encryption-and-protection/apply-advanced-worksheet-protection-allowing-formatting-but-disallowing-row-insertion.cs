// Title: Aspose.Cells .NET – Protect Worksheet, Allow Formatting, Block Row Insertion (Password Protected)
// Description: Creates a new workbook, accesses the first worksheet, enables cell/column/row formatting, disables row insertion, sets a password, applies full protection with ProtectionType.All, and saves the file as AdvancedProtection.xlsx.
// Keywords: Aspose.Cells | worksheet protection | allow formatting | block row insertion | password protected | C# | .NET | Protect method | ProtectionType.All
// Common Searches: Aspose.Cells allow formatting but prevent row insertion | protect worksheet with password in Aspose.Cells .NET | disable row insertion while enabling formatting Aspose.Cells | C# Aspose.Cells worksheet protection settings | how to set ProtectionType.All in Aspose.Cells
// Developer Intent: The developer needs to secure a worksheet so users can format cells, columns, and rows but cannot add new rows, using a password‑protected protection scheme.
// Use Cases: Distribute a template that lets collaborators style data without altering the row layout. | Publish a financial report where visual tweaks are allowed but the row count must stay fixed. | Share a spreadsheet for team review, permitting formatting changes while preventing structural modifications.
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet, enable formatting, disable row insertion, and apply a password. | Show how to modify the protection object to also block column insertion while keeping formatting permissions active. | Explain the differences between ProtectionType.All, Objects, and Scenarios and when to use each in Aspose.Cells .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Creates a new workbook, accesses the first worksheet, enables cell/column/row formatting, disables row insertion, sets a password, applies full protection with ProtectionType.All, and saves the file as AdvancedProtection.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the protection object of the worksheet
            Protection protection = sheet.Protection;

            // Allow formatting of cells, columns and rows
            protection.AllowFormattingCell = true;
            protection.AllowFormattingColumn = true;
            protection.AllowFormattingRow = true;

            // Disallow insertion of rows (default is false, set explicitly for clarity)
            protection.AllowInsertingRow = false;

            // Optional: set a password for the protection
            protection.Password = "MySecretPassword";

            // Apply protection to the worksheet (protect all aspects)
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("AdvancedProtection.xlsx");
        }
    }
}
