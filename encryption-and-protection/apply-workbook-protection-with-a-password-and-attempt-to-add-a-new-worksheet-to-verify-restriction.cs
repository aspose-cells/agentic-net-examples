// Title: How to protect an Excel workbook's structure with a password and verify that adding a new worksheet is blocked using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect workbook structure with password and block new sheets | C# code to apply password protection to Excel workbook structure using Aspose.Cells | How to test workbook structure protection by trying to add a worksheet in Aspose.Cells | Exception thrown when adding worksheet to password‑protected workbook in Aspose.Cells .NET | Save an Excel file after applying structure protection with Aspose.Cells for .NET
// Tags: protect workbook structure Aspose.Cells | password protected Excel workbook .NET | block worksheet addition Aspose.Cells | workbook protection exception handling C# | save protected workbook Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, applies structure protection with the password "MySecretPwd", attempts to add a worksheet (which fails and throws an exception due to the protection), catches and logs the error, and finally saves the workbook as a password‑protected Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Apply workbook protection (structure) with a password
        workbook.Protect(ProtectionType.Structure, "MySecretPwd");

        // Attempt to add a new worksheet; this should be blocked by the protection
        try
        {
            workbook.Worksheets.Add("NewSheet");
            Console.WriteLine("Worksheet added successfully (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to add worksheet as expected: " + ex.Message);
        }

        // Save the protected workbook
        workbook.Save("ProtectedWorkbook.xlsx");
    }
}
