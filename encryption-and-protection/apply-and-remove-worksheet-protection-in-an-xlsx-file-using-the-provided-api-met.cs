using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and protect its first worksheet with a password
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];               // get the first worksheet

            // Protect the worksheet with all protection types and a password
            sheet.Protect(ProtectionType.All, "mySecretPwd", null);

            // Save the protected workbook
            string protectedPath = "ProtectedWorksheet.xlsx";
            workbook.Save(protectedPath);
            Console.WriteLine($"Worksheet protected and saved to '{protectedPath}'.");

            // -----------------------------------------------------------------
            // 2. Load the protected workbook, verify protection, then unprotect it
            // -----------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(protectedPath);  // load the saved file
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Verify that the worksheet is indeed protected
            Console.WriteLine($"Is worksheet protected? {loadedSheet.IsProtected}");

            // Unprotect the worksheet using the correct password
            loadedSheet.Unprotect("mySecretPwd");

            // Verify that the worksheet is now unprotected
            Console.WriteLine($"Is worksheet protected after unprotecting? {loadedSheet.IsProtected}");

            // Save the unprotected workbook
            string unprotectedPath = "UnprotectedWorksheet.xlsx";
            loadedWorkbook.Save(unprotectedPath);
            Console.WriteLine($"Worksheet unprotected and saved to '{unprotectedPath}'.");
        }
    }
}