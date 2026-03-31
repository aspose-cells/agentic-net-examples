using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default has one worksheet)
            Workbook workbook = new Workbook();

            // Protect only the workbook structure with a password
            // ProtectionType.Structure protects the ability to add, delete, rename, hide, or unhide worksheets
            workbook.Protect(ProtectionType.Structure, "mySecretPwd");

            // Save the protected workbook
            string protectedPath = "ProtectedWorkbook.xlsx";
            workbook.Save(protectedPath, SaveFormat.Xlsx);
            workbook.Dispose();

            // Load the protected workbook
            Workbook loadedWorkbook = new Workbook(protectedPath);

            // Verify that the workbook is protected
            Console.WriteLine("Is workbook protected after load? " + loadedWorkbook.Settings.IsProtected);
            Console.WriteLine("Is workbook protected with password? " + loadedWorkbook.IsWorkbookProtectedWithPassword);

            // Unprotect the workbook using the same password
            loadedWorkbook.Unprotect("mySecretPwd");

            // Verify that the protection has been removed
            Console.WriteLine("Is workbook protected after unprotect? " + loadedWorkbook.Settings.IsProtected);

            // Save the unprotected workbook
            string unprotectedPath = "UnprotectedWorkbook.xlsx";
            loadedWorkbook.Save(unprotectedPath, SaveFormat.Xlsx);
            loadedWorkbook.Dispose();
        }
    }
}