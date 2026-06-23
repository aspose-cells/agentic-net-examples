using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with an initial password
        sheet.Protect(ProtectionType.All, "Password123", null);
        Console.WriteLine("Initially protected: " + sheet.IsProtected);

        // Remove protection using the correct password
        sheet.Unprotect("Password123");
        Console.WriteLine("After unprotect: " + !sheet.IsProtected);

        // Re‑apply protection with a case‑sensitive password
        sheet.Protect(ProtectionType.All, "CaseSensitivePass", null);
        Console.WriteLine("Re‑protected with case‑sensitive password: " + sheet.IsProtected);

        // Save the workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}