// Title: Handle Wrong Password Exception When Unprotecting an Aspose.Cells Worksheet (C#)
// Description: Demonstrates protecting a worksheet with a known password, attempting to unprotect it using an incorrect password, and capturing the resulting exception in a try‑catch block. The example also shows saving the workbook.
// Keywords: Aspose.Cells unprotect worksheet | incorrect password exception | Worksheet.Unprotect try catch | C# Aspose.Cells protection error handling | catch Aspose.Cells unprotect error
// Common Searches: Aspose.Cells unprotect worksheet with wrong password | exception thrown by Worksheet.Unprotect when password is invalid | how to catch unprotect error in Aspose.Cells C# | protect and unprotect sheet Aspose.Cells example
// Developer Intent: Show how to attempt to unprotect a protected worksheet using an invalid password and retrieve the exception message without crashing the application.
// Use Cases: Validate user‑entered passwords before calling Unprotect and log failures. | Prevent runtime crashes by wrapping Worksheet.Unprotect in a try‑catch block. | Display a user‑friendly error message when a wrong password is supplied for a protected sheet.
// AI Prompts: Write C# code that protects an Aspose.Cells worksheet, then tries to unprotect it with an incorrect password and returns the caught exception message. | Explain which exception type Worksheet.Unprotect throws on a wrong password and suggest best practices for handling it in .NET.

using System;
using Aspose.Cells;

// Demonstrates protecting a worksheet with a known password, attempting to unprotect it using an incorrect password, and capturing the resulting exception in a try‑catch block. The example also shows saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Protect the worksheet with a known password
        worksheet.Protect(ProtectionType.All, "correctPassword", null);
        Console.WriteLine("Worksheet is protected: " + worksheet.IsProtected);

        try
        {
            // Attempt to unprotect the worksheet using an incorrect password
            worksheet.Unprotect("wrongPassword");
            Console.WriteLine("Worksheet unprotected: " + !worksheet.IsProtected);
        }
        catch (Exception ex)
        {
            // Capture and display the exception thrown due to incorrect password
            Console.WriteLine("Exception caught while unprotecting: " + ex.Message);
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("UnprotectDemo.xlsx");
    }
}
