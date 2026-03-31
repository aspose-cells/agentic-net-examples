using System;
using Aspose.Cells;

namespace AsposeCellsStructuralProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------- Create and protect workbook -------------------
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Protected Structure Demo");

            // Protect the workbook structure with a password
            // ProtectionType.Structure protects the ability to add, delete, rename, hide, or unhide worksheets
            workbook.Protect(ProtectionType.Structure, "MySecretPwd");

            // Save the protected workbook
            workbook.Save("ProtectedWorkbook.xlsx");

            // ------------------- Load and verify protection -------------------
            // Load the previously saved workbook
            Workbook protectedWb = new Workbook("ProtectedWorkbook.xlsx");

            // Verify that the workbook is protected with a password
            Console.WriteLine("Is workbook protected with password? " + protectedWb.IsWorkbookProtectedWithPassword);

            // ------------------- Unprotect and save unprotected version -------------------
            // Unprotect the workbook using the same password
            protectedWb.Unprotect("MySecretPwd");

            // Verify that protection has been removed
            Console.WriteLine("Is workbook still protected? " + protectedWb.IsWorkbookProtectedWithPassword);

            // Save the unprotected workbook
            protectedWb.Save("UnprotectedWorkbook.xlsx");
        }
    }
}