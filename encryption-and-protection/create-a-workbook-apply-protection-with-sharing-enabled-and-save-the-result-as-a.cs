using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Enable sharing for the workbook
        wb.Settings.Shared = true;

        // Protect the shared workbook with a password
        wb.ProtectSharedWorkbook("myPassword");

        // Save the protected, shared workbook as XLSX
        wb.Save("ProtectedSharedWorkbook.xlsx");
    }
}