using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SharedWorkbookProtectionDemo
    {
        public static void Run()
        {
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Shared Workbook Demo");

            // Protect the shared workbook with a password
            string password = "mySecretPwd";
            wb.ProtectSharedWorkbook(password);

            // Save the protected shared workbook
            string protectedPath = "ProtectedSharedWorkbook.xlsx";
            wb.Save(protectedPath);

            // Load the protected workbook using the file path
            Workbook wbProtected = new Workbook(protectedPath);

            // Verify that the workbook is protected (shared protection)
            Console.WriteLine("Workbook is protected (shared): " + wbProtected.Settings.IsProtected);

            // Unprotect the shared workbook using the same password
            wbProtected.UnprotectSharedWorkbook(password);

            // Save the unprotected workbook
            string unprotectedPath = "UnprotectedSharedWorkbook.xlsx";
            wbProtected.Save(unprotectedPath);

            // Load the unprotected workbook to confirm removal of protection
            Workbook wbFinal = new Workbook(unprotectedPath);
            Console.WriteLine("Workbook is protected after unprotect: " + wbFinal.Settings.IsProtected);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SharedWorkbookProtectionDemo.Run();
        }
    }
}