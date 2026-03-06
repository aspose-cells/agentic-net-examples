using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectUnprotectWorkbookDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Protect the workbook (structure and windows) with a password
            wb.Protect(ProtectionType.All, "myPassword");

            // Save the protected workbook
            wb.Save("ProtectedWorkbook.xlsx");

            // Load the protected workbook
            Workbook protectedWb = new Workbook("ProtectedWorkbook.xlsx");

            // Verify protection status
            Console.WriteLine("Workbook protected (Settings.IsProtected): " + protectedWb.Settings.IsProtected);
            Console.WriteLine("Workbook protected with password: " + protectedWb.IsWorkbookProtectedWithPassword);

            // Unprotect the workbook using the correct password
            protectedWb.Unprotect("myPassword");

            // Save the unprotected workbook
            protectedWb.Save("UnprotectedWorkbook.xlsx");

            // ----- Shared workbook protection example -----
            // Create another workbook
            Workbook sharedWb = new Workbook();

            // Protect the shared workbook with a password
            sharedWb.ProtectSharedWorkbook("sharedPwd");

            // Save the protected shared workbook
            sharedWb.Save("ProtectedSharedWorkbook.xlsx");

            // Load the protected shared workbook
            Workbook loadedShared = new Workbook("ProtectedSharedWorkbook.xlsx");

            // Verify shared workbook protection status
            Console.WriteLine("Shared workbook protected (Settings.IsProtected): " + loadedShared.Settings.IsProtected);

            // Unprotect the shared workbook
            loadedShared.UnprotectSharedWorkbook("sharedPwd");

            // Save the unprotected shared workbook
            loadedShared.Save("UnprotectedSharedWorkbook.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectUnprotectWorkbookDemo.Run();
        }
    }
}