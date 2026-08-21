// Title: C# – Protect an Excel worksheet with a password and simulate expiration after a delay using Aspose.Cells
// Description: Demonstrates how to create a workbook, apply full worksheet protection with a password, verify that an incorrect password fails, wait a configurable interval, then remove the protection with the correct password, check the IsProtected flag, and save the file in an unprotected state.
// Keywords: Aspose.Cells C# worksheet protection | Excel password protection .NET | protect worksheet with password Aspose.Cells | worksheet protection expiration simulation | unprotect worksheet after delay | IsProtected property Aspose.Cells | Excel security timeout C# | Aspose.Cells protect/unprotect example
// Common Searches: protect Excel worksheet with password using Aspose.Cells C# | remove worksheet protection after a timeout Aspose.Cells | how to test worksheet protection expiration in .NET | Aspose.Cells unprotect worksheet with correct password | simulate protection expiry for Excel file
// Developer Intent: The developer needs to apply password‑based protection to a worksheet, confirm that wrong passwords are rejected, wait a set period, then automatically lift the protection to verify expiration behavior.
// Use Cases: Secure a worksheet before distribution and make it editable after a predefined interval. | Validate that only the correct password can unprotect the sheet while an incorrect one throws an exception. | Programmatically check the protection status and save the workbook once the protection is lifted.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet using a password and automatically unprotects it after N seconds. | Explain how to catch exceptions when an invalid password is used to unprotect a worksheet in Aspose.Cells. | Show how to query the IsProtected property, wait for a timeout, and save the workbook in an unprotected state.

using System;
using System.Threading;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply full worksheet protection with a password, verify that an incorrect password fails, wait a configurable interval, then remove the protection with the correct password, check the IsProtected flag, and save the file in an unprotected state.
class WorksheetProtectionExpirationDemo
{
    static void Main()
    {
        // Create a new workbook and protect the first worksheet with a password
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        string password = "secret123";

        // Protect the worksheet (all protection types) with the password
        worksheet.Protect(ProtectionType.All, password, null);
        workbook.Save("ProtectedWorksheet.xlsx");
        Console.WriteLine("Worksheet protected with password.");

        // Load the protected workbook
        Workbook loadedWorkbook = new Workbook("ProtectedWorksheet.xlsx");
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

        // Attempt to unprotect with an incorrect password (should fail)
        try
        {
            loadedWorksheet.Unprotect("wrongPassword");
            Console.WriteLine("Unexpectedly unprotected with wrong password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to unprotect with wrong password: " + ex.Message);
        }

        // Define expiration interval (in seconds)
        int expirationInterval = 5;
        Console.WriteLine($"Waiting {expirationInterval} seconds for protection to expire...");
        Thread.Sleep(expirationInterval * 1000);

        // Simulate expiration by removing protection using the correct password
        loadedWorksheet.Unprotect(password);
        Console.WriteLine("Protection expired; worksheet is now unprotected.");

        // Verify that the worksheet is no longer protected
        Console.WriteLine("Worksheet IsProtected: " + loadedWorksheet.IsProtected);

        // Save the workbook after expiration
        loadedWorkbook.Save("UnprotectedWorksheet.xlsx");
    }
}
