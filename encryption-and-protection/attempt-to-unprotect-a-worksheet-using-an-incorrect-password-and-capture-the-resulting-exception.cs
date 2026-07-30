// Title: Handle wrong‑password exception when unprotecting an Aspose.Cells worksheet (C#)
// Description: Demonstrates how to protect a worksheet, attempt to unprotect it with an incorrect password, catch the resulting exception, verify the worksheet stays protected, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Worksheet.Unprotect | wrong password | exception handling | IsProtected | protect worksheet | save workbook
// Common Searches: Aspose.Cells unprotect worksheet with wrong password | catch exception Worksheet.Unprotect C# | check IsProtected after failed unprotect Aspose.Cells | how to handle wrong password when unprotecting a worksheet | save protected workbook after exception Aspose.Cells
// Developer Intent: The developer wants to try unprotecting a protected worksheet using an incorrect password, capture the thrown exception, and confirm the worksheet remains protected.
// Use Cases: Log detailed error information when an unprotect operation fails due to a bad password. | Validate password input before calling Unprotect to avoid exceptions in production code. | Ensure the worksheet's protection state (IsProtected) is unchanged after a failed unprotect attempt before saving the file.
// AI Prompts: Generate a C# method that attempts Worksheet.Unprotect with a given password, returns true on success, and logs the exception message while returning false on failure. | Create sample code that protects a worksheet, tries to unprotect it with an invalid password inside a try‑catch block, prints the exception, verifies worksheet.IsProtected is still true, and then saves the workbook.

using System;
using Aspose.Cells;

// Demonstrates how to protect a worksheet, attempt to unprotect it with an incorrect password, catch the resulting exception, verify the worksheet stays protected, and save the workbook using Aspose.Cells for .NET.
public class UnprotectWorksheetDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Protect the worksheet with a password
        worksheet.Protect(ProtectionType.All, "correctPassword", null);
        Console.WriteLine("Worksheet is protected: " + worksheet.IsProtected);

        // Attempt to unprotect using an incorrect password and capture the exception
        try
        {
            worksheet.Unprotect("wrongPassword");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Caught exception when unprotecting with wrong password: " + ex.Message);
        }

        // Verify that the worksheet remains protected
        Console.WriteLine("Worksheet still protected after failed unprotect: " + worksheet.IsProtected);

        // Save the workbook (optional, demonstrates save lifecycle)
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
