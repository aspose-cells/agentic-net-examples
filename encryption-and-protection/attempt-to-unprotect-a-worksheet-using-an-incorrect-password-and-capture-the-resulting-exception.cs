// Title: Handling the exception thrown by Aspose.Cells when unprotecting a worksheet with an invalid password in C#
// Common Searches: Aspose.Cells C# catch exception for invalid worksheet unprotect password | How to detect wrong password error when calling worksheet.Unprotect in Aspose.Cells | What exception does Aspose.Cells throw if worksheet unprotect password is incorrect | C# example of handling password mismatch during worksheet unprotect with Aspose.Cells | Unprotect protected worksheet using Aspose.Cells and capture error in .NET
// Tags: worksheet unprotect exception handling Aspose.Cells | Aspose.Cells worksheet protection password mismatch | C# catch Aspose.Cells protection error | Aspose.Cells unprotect password error handling | Aspose.Cells worksheet protection API error handling

using System;
using Aspose.Cells;

// The sample creates a new Workbook, protects its first worksheet with a password, then attempts to unprotect it using an incorrect password, catches the resulting exception, and writes the exception message to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Protect the worksheet with a password
            worksheet.Protect(ProtectionType.All);
            worksheet.Protection.Password = "correctPassword";

            try
            {
                // Attempt to unprotect using an incorrect password
                worksheet.Unprotect("wrongPassword");
            }
            catch (Exception ex)
            {
                // Capture and display the exception details
                Console.WriteLine("Exception caught: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
