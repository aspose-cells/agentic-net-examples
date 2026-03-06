using System;
using Aspose.Cells;

class WorkbookProtectionDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword");

        // Save the protected workbook to disk
        workbook.Save("protected_workbook.xlsx", SaveFormat.Xlsx);

        // Load the previously saved protected workbook
        Workbook protectedWb = new Workbook("protected_workbook.xlsx");

        // Check if the workbook is protected with a password
        Console.WriteLine("Is workbook protected with password? " + protectedWb.IsWorkbookProtectedWithPassword);

        // Unprotect the workbook using the correct password
        protectedWb.Unprotect("myPassword");

        // Save the unprotected workbook
        protectedWb.Save("unprotected_workbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
        protectedWb.Dispose();
    }
}