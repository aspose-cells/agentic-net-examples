// Title: Catch the exception from an invalid worksheet Unprotect call with Aspose.Cells for .NET
// Description: Demonstrates how to protect a worksheet, attempt to unprotect it with an incorrect password, capture the thrown exception, display the error message, and save the workbook using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | worksheet protection | Unprotect wrong password | exception handling | IsProtected | ProtectionType.All | catch invalid password | Aspose.Cells error message
// Common Searches: Aspose.Cells catch exception invalid worksheet password | Unprotect worksheet with wrong password C# Aspose.Cells | What exception is thrown by Aspose.Cells Unprotect when password is incorrect | How to handle failed worksheet unprotect in Aspose.Cells
// Developer Intent: Show how to detect and handle the error that occurs when Unprotect is called with an incorrect password.
// Use Cases: Validate user‑entered passwords before calling Unprotect and log failures. | Prevent application crashes by catching the exception and showing a friendly message. | Record security audit entries whenever an unauthorized unprotect attempt is made.
// AI Prompts: Generate C# code using Aspose.Cells that protects a worksheet, tries to unprotect it with a bad password, and logs the exception message. | Explain which exception type Aspose.Cells throws on a failed Unprotect call and how to extract detailed information. | Provide a robust pattern for checking IsProtected after a failed unprotect operation and recording the incident.

using System;
using Aspose.Cells;

// Demonstrates how to protect a worksheet, attempt to unprotect it with an incorrect password, capture the thrown exception, display the error message, and save the workbook using Aspose.Cells in C#.
class UnprotectWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with a known password
        sheet.Protect(ProtectionType.All, "correctPassword", null);
        Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

        // Try to unprotect using an incorrect password and capture the exception
        try
        {
            sheet.Unprotect("wrongPassword");
            Console.WriteLine("Worksheet unprotected (unexpected): " + !sheet.IsProtected);
        }
        catch (Exception ex)
        {
            // Expected exception for incorrect password
            Console.WriteLine("Exception caught while unprotecting with wrong password: " + ex.Message);
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("UnprotectDemo.xlsx");
    }
}
