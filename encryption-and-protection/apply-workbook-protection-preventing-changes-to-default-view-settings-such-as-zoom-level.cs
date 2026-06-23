using Aspose.Cells;

class WorkbookProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook's window (default view settings such as zoom level) with a password
        workbook.Protect(ProtectionType.Windows, "MySecurePassword");

        // Save the protected workbook
        workbook.Save("ProtectedViewSettings.xlsx");
    }
}