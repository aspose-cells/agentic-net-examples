using System;
using Aspose.Cells;

namespace AsposeCellsSharedWorkbookProtection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the original workbook (must be a shared workbook)
            string inputPath = "SharedWorkbook.xlsx";

            // Path for the protected workbook
            string protectedPath = "ProtectedSharedWorkbook.xlsx";

            // Path for the unprotected workbook
            string unprotectedPath = "UnprotectedSharedWorkbook.xlsx";

            // Password to use for protection
            string password = "mySecretPwd";

            // -------------------------------------------------
            // Load the existing shared workbook
            // -------------------------------------------------
            Workbook wb = new Workbook(inputPath);

            // -------------------------------------------------
            // Protect the shared workbook with a password
            // -------------------------------------------------
            wb.ProtectSharedWorkbook(password);

            // Save the protected workbook
            wb.Save(protectedPath);

            // -------------------------------------------------
            // Load the protected workbook to verify and unprotect
            // -------------------------------------------------
            Workbook wbProtected = new Workbook(protectedPath);

            // Unprotect the shared workbook using the same password
            wbProtected.UnprotectSharedWorkbook(password);

            // Save the unprotected workbook
            wbProtected.Save(unprotectedPath);

            // -------------------------------------------------
            // Optional: display protection status for verification
            // -------------------------------------------------
            Console.WriteLine("Original workbook protected? " + wb.Settings.IsProtected);
            Console.WriteLine("Protected workbook protected? " + wbProtected.Settings.IsProtected);
            Console.WriteLine("Unprotected workbook protected? " + new Workbook(unprotectedPath).Settings.IsProtected);
        }
    }
}