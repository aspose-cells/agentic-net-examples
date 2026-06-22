using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet is unprotected before adding a protected range
        worksheet.Unprotect();

        // Add a protected range (cells A1:B2) and set a password for the range
        int rangeIndex = worksheet.AllowEditRanges.Add("MyProtectedRange", 0, 0, 1, 1);
        ProtectedRange protectedRange = worksheet.AllowEditRanges[rangeIndex];
        protectedRange.Password = "rangePassword";

        // Protect the worksheet with its own password
        worksheet.Protect(ProtectionType.All, "sheetPassword", null);

        // Save the workbook
        string fileName = "ProtectedRangeDemo.xlsx";
        workbook.Save(fileName);

        // Load the saved workbook to verify that the protected range remains password‑protected
        Workbook verifyWorkbook = new Workbook(fileName);
        ProtectedRange verifyRange = verifyWorkbook.Worksheets[0].AllowEditRanges[0];

        // Output verification result
        Console.WriteLine("Protected range is password protected: " + verifyRange.IsProtectedWithPassword);
    }
}