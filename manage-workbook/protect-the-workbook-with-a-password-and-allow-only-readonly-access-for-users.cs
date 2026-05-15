using System;
using Aspose.Cells;

class ProtectWorkbookReadOnly
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set write‑protection password (required to modify the file)
        workbook.Settings.WriteProtection.Password = "ownerPassword";

        // Recommend opening the file as read‑only for users who do not know the password
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Optional: set the author of the protection
        workbook.Settings.WriteProtection.Author = "Admin";

        // Save the protected workbook
        workbook.Save("ProtectedReadOnlyWorkbook.xlsx");

        // The workbook is now write‑protected; users can open it in read‑only mode
        // unless they provide the password to modify it.
    }
}