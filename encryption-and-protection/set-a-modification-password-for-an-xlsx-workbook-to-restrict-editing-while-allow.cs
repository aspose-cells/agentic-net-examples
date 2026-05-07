using System;
using Aspose.Cells;

class SetModificationPassword
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set a write‑protection password to restrict editing while allowing view access
        workbook.Settings.WriteProtection.Password = "modify123";

        // Optionally suggest opening the file as read‑only
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook in XLSX format
        workbook.Save("ModifiedProtectedWorkbook.xlsx");

        // Load the saved workbook to verify that write protection is applied
        Workbook loaded = new Workbook("ModifiedProtectedWorkbook.xlsx");
        bool isWriteProtected = loaded.Settings.WriteProtection.IsWriteProtected;
        Console.WriteLine("Is workbook write‑protected? " + isWriteProtected);
    }
}