using Aspose.Cells;
using System;

class UnprotectWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with a password
        sheet.Protect(ProtectionType.All, "mySecret", null);
        Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

        // Verify the password (optional, demonstrates verification)
        bool passwordCorrect = sheet.Protection.VerifyPassword("mySecret");
        Console.WriteLine("Password verification result: " + passwordCorrect);

        // Unprotect the worksheet using the correct password
        sheet.Unprotect("mySecret");
        Console.WriteLine("Worksheet protected after unprotect: " + sheet.IsProtected);

        // Verify that editing is now permitted (IsProtected should be false)
        if (!sheet.IsProtected)
        {
            Console.WriteLine("Editing is now permitted.");
        }

        // Save the workbook
        workbook.Save("UnprotectedWorksheet.xlsx");
    }
}