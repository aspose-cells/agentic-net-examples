using System;
using Aspose.Cells;

class ProtectUnprotectWorksheet
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with a password and all protection types
        sheet.Protect(ProtectionType.All, "myPassword", null);

        // Save the protected workbook
        workbook.Save("ProtectedWorksheet.xlsx");

        // Load the protected workbook
        Workbook loaded = new Workbook("ProtectedWorksheet.xlsx");

        // Check if the worksheet is protected
        Console.WriteLine("Worksheet protected: " + loaded.Worksheets[0].IsProtected);

        // Unprotect the worksheet using the correct password
        loaded.Worksheets[0].Unprotect("myPassword");

        // Verify that the worksheet is now unprotected
        Console.WriteLine("Worksheet protected after unprotect: " + loaded.Worksheets[0].IsProtected);

        // Save the unprotected workbook
        loaded.Save("UnprotectedWorksheet.xlsx");
    }
}