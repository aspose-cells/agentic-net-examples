using Aspose.Cells;

class ProtectWorkbookReadOnly
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the password required to modify the file
        wb.Settings.WriteProtection.Password = "modifyPwd";

        // Recommend opening the file as read‑only
        wb.Settings.WriteProtection.RecommendReadOnly = true;

        // Optional: set the author of the protection
        wb.Settings.WriteProtection.Author = "Admin";

        // Save the protected workbook
        wb.Save("ReadOnlyProtectedWorkbook.xlsx");
    }
}