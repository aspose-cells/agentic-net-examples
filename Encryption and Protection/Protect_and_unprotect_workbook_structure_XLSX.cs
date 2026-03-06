using System;
using Aspose.Cells;

class ProtectUnprotectWorkbook
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword");

        // Check protection status
        Console.WriteLine("Workbook protected? " + workbook.Settings.IsProtected);
        Console.WriteLine("Protected with password? " + workbook.IsWorkbookProtectedWithPassword);

        // Save the protected workbook
        workbook.Save("protected_workbook.xlsx", SaveFormat.Xlsx);

        // Load the protected workbook
        Workbook loaded = new Workbook("protected_workbook.xlsx");

        // Verify that the loaded workbook is still protected
        Console.WriteLine("Loaded workbook protected? " + loaded.Settings.IsProtected);

        // Unprotect the workbook using the password
        loaded.Unprotect("myPassword");

        // Verify that the workbook is now unprotected
        Console.WriteLine("After unprotect, protected? " + loaded.Settings.IsProtected);

        // Save the unprotected workbook
        loaded.Save("unprotected_workbook.xlsx", SaveFormat.Xlsx);
    }
}