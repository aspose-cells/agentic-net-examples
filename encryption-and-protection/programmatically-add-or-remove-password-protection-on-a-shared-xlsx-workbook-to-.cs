using System;
using Aspose.Cells;

namespace AsposeCellsPasswordProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------- Create a new workbook -------------------
            // The Workbook constructor is the creation rule.
            Workbook wb = new Workbook();

            // Add some sample data to the first worksheet.
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // ------------------- Protect the shared workbook -------------------
            // ProtectSharedWorkbook adds password protection to a shared workbook.
            string password = "MySecretPwd";
            wb.ProtectSharedWorkbook(password);

            // Save the protected workbook.
            // The Save method is the saving rule.
            string protectedPath = "ProtectedSharedWorkbook.xlsx";
            wb.Save(protectedPath);

            // ------------------- Load the protected workbook -------------------
            // The Workbook(string) constructor is the loading rule.
            Workbook protectedWb = new Workbook(protectedPath);

            // Verify that the workbook is protected.
            // Settings.IsProtected indicates structure/window protection.
            Console.WriteLine("Workbook is protected (Settings.IsProtected): " + protectedWb.Settings.IsProtected);
            // IsWorkbookProtectedWithPassword indicates password protection.
            Console.WriteLine("Workbook is password protected: " + protectedWb.IsWorkbookProtectedWithPassword);

            // ------------------- Unprotect the shared workbook -------------------
            // Remove the password protection using the same password.
            protectedWb.UnprotectSharedWorkbook(password);

            // Save the unprotected workbook.
            string unprotectedPath = "UnprotectedSharedWorkbook.xlsx";
            protectedWb.Save(unprotectedPath);

            // Load the unprotected workbook to confirm removal of protection.
            Workbook unprotectedWb = new Workbook(unprotectedPath);
            Console.WriteLine("After unprotecting, Settings.IsProtected: " + unprotectedWb.Settings.IsProtected);
            Console.WriteLine("After unprotecting, IsWorkbookProtectedWithPassword: " + unprotectedWb.IsWorkbookProtectedWithPassword);
        }
    }
}