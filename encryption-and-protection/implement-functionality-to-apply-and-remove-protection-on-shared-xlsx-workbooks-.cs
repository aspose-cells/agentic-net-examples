using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    public class SharedWorkbookProtection
    {
        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook wb = new Workbook();

            // Apply shared workbook protection with a password
            wb.ProtectSharedWorkbook("mySecretPwd");

            // Save the protected workbook to disk
            wb.Save("ProtectedSharedWorkbook.xlsx");

            // Load the protected workbook from disk
            Workbook protectedWb = new Workbook("ProtectedSharedWorkbook.xlsx");

            // Verify that the workbook is protected (Settings.IsProtected)
            Console.WriteLine("Workbook is protected after load: " + protectedWb.Settings.IsProtected);

            // Remove protection using the same password
            protectedWb.UnprotectSharedWorkbook("mySecretPwd");

            // Save the unprotected workbook
            protectedWb.Save("UnprotectedSharedWorkbook.xlsx");

            // Load the unprotected workbook to confirm removal of protection
            Workbook unprotectedWb = new Workbook("UnprotectedSharedWorkbook.xlsx");
            Console.WriteLine("Workbook is protected after unprotect: " + unprotectedWb.Settings.IsProtected);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SharedWorkbookProtection.Run();
        }
    }
}