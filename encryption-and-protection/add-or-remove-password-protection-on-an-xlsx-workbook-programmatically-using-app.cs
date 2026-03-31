using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create and protect workbook --------------------
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add some sample data
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Set file encryption password (protects the whole file)
            wb.Settings.Password = "FileEncryptPwd";

            // Optionally protect workbook structure with a password
            wb.Protect(ProtectionType.All, "StructurePwd");

            // Save the password‑protected workbook
            string protectedPath = "protected.xlsx";
            wb.Save(protectedPath);
            Console.WriteLine($"Workbook saved with encryption and structure protection: {protectedPath}");

            // -------------------- Load, verify and remove protection --------------------
            // Load the protected workbook using the encryption password
            LoadOptions loadOpts = new LoadOptions();
            loadOpts.Password = "FileEncryptPwd";
            Workbook loadedWb = new Workbook(protectedPath, loadOpts);

            // Verify that the workbook is still protected
            Console.WriteLine($"Is workbook protected with password? {loadedWb.IsWorkbookProtectedWithPassword}");
            Console.WriteLine($"Is workbook write‑protected? {loadedWb.Settings.WriteProtection.IsWriteProtected}");

            // Remove structure protection
            loadedWb.Unprotect("StructurePwd");

            // Remove file encryption by clearing the password
            loadedWb.Settings.Password = null;

            // Save the unprotected workbook
            string unprotectedPath = "unprotected.xlsx";
            loadedWb.Save(unprotectedPath);
            Console.WriteLine($"Workbook saved without any password protection: {unprotectedPath}");
        }
    }
}