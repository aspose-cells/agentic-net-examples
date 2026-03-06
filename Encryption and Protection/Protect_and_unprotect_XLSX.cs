using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // -------------------------------------------------
            // 1. Protect a worksheet with a password and save
            // -------------------------------------------------
            Workbook wb1 = new Workbook();                         // create new workbook
            Worksheet sheet1 = wb1.Worksheets[0];                  // get first worksheet
            sheet1.Protect(ProtectionType.All, "sheetPwd", null); // protect worksheet with password
            wb1.Save("WorksheetProtected.xlsx");                  // save the protected worksheet

            // -------------------------------------------------
            // 2. Load the protected worksheet, unprotect it, and save
            // -------------------------------------------------
            Workbook wb2 = new Workbook("WorksheetProtected.xlsx"); // load workbook
            Worksheet sheet2 = wb2.Worksheets[0];
            sheet2.Unprotect("sheetPwd");                         // unprotect using correct password
            wb2.Save("WorksheetUnprotected.xlsx");                // save the unprotected workbook

            // -------------------------------------------------
            // 3. Protect the entire workbook structure with a password and save
            // -------------------------------------------------
            Workbook wb3 = new Workbook();                         // create new workbook
            wb3.Protect(ProtectionType.Structure, "workbookPwd"); // protect workbook structure
            wb3.Save("WorkbookProtected.xlsx");                    // save protected workbook

            // -------------------------------------------------
            // 4. Load the protected workbook, unprotect it, and save
            // -------------------------------------------------
            Workbook wb4 = new Workbook("WorkbookProtected.xlsx"); // load protected workbook
            wb4.Unprotect("workbookPwd");                         // unprotect using correct password
            wb4.Save("WorkbookUnprotected.xlsx");                  // save unprotected workbook

            // -------------------------------------------------
            // 5. Apply write protection to a workbook, validate password, and save
            // -------------------------------------------------
            Workbook wb5 = new Workbook();                         // create new workbook
            wb5.Settings.WriteProtection.Password = "writePwd";   // set write‑protection password
            bool isWriteProtected = wb5.Settings.WriteProtection.IsWriteProtected;
            Console.WriteLine("Is write protected: " + isWriteProtected);
            bool isValid = wb5.Settings.WriteProtection.ValidatePassword("writePwd");
            Console.WriteLine("Password validation result: " + isValid);
            wb5.Save("WriteProtected.xlsx");                      // save write‑protected workbook

            // -------------------------------------------------
            // 6. Load the write‑protected workbook and verify protection status
            // -------------------------------------------------
            Workbook wb6 = new Workbook("WriteProtected.xlsx");   // load workbook
            bool isWriteProtectedLoaded = wb6.Settings.WriteProtection.IsWriteProtected;
            Console.WriteLine("Loaded workbook write protected: " + isWriteProtectedLoaded);
        }
    }
}