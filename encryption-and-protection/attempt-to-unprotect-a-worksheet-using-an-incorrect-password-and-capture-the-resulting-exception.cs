// Title: C# code sample: Catch exception when unprotecting an Aspose.Cells worksheet with an invalid password
// Common Searches: how to catch exception for invalid password when unprotecting a worksheet using Aspose.Cells in C# | Aspose.Cells unprotect worksheet throws error with wrong password example C# | C# handling of Aspose.Cells worksheet protection failure due to incorrect password | capture Aspose.Cells worksheet unprotect exception caused by mismatched password
// Tags: Aspose.Cells worksheet unprotect API C# | invalid password handling Aspose.Cells protection | exception handling for worksheet unprotect Aspose.Cells | Aspose.Cells Protect and Unprotect methods usage | C# capture Aspose.Cells protection error | worksheet password mismatch Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, protects the first worksheet with a known password, then attempts to unprotect it using an incorrect password. The resulting exception is caught and its message is written to the console, demonstrating proper error handling for worksheet protection failures.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Protect the worksheet with a known password (protect all options)
            // The third parameter is the old password; for a new protection it can be empty
            sheet.Protect(ProtectionType.All, "correctPassword", string.Empty);

            try
            {
                // Attempt to unprotect the worksheet using an incorrect password
                sheet.Unprotect("wrongPassword");
            }
            catch (Exception ex)
            {
                // Capture and display the exception that occurs
                Console.WriteLine("Exception caught while unprotecting worksheet:");
                Console.WriteLine(ex.Message);
            }

            // (Optional) Save the workbook if needed
            // workbook.Save("ProtectedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine("An unexpected error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
