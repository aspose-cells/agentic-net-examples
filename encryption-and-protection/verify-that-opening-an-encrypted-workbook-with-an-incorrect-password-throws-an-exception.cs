// Title: Aspose.Cells for .NET – Verify exception when opening an encrypted workbook with a wrong password (C#)
// Description: Shows how to encrypt a workbook, save it, and then load it with an incorrect password using LoadOptions. The code catches the expected exception, proving that Aspose.Cells rejects invalid passwords.
// Keywords: Aspose.Cells | .NET | C# | encrypted workbook | password protection | LoadOptions | wrong password | exception handling | Workbook security
// Common Searches: Aspose.Cells load encrypted workbook wrong password | C# catch exception incorrect password Aspose.Cells | verify password protection error Aspose.Cells .NET | how to test invalid password with Aspose.Cells
// Developer Intent: Confirm that loading a password‑protected workbook with an invalid password throws an exception in Aspose.Cells.
// Use Cases: Unit test to ensure the library blocks access when the password is incorrect. | Automated validation of encrypted Excel files before batch processing. | Graceful error handling in services that accept user‑uploaded, password‑protected workbooks.
// AI Prompts: Generate an NUnit test that asserts Aspose.Cells throws the correct exception when a workbook encrypted with a known password is opened with a wrong password. | Provide a try‑catch example that logs the specific Aspose.Cells exception type for an incorrect password scenario. | Create code that distinguishes between a missing password and an incorrect password when loading a protected workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to encrypt a workbook, save it, and then load it with an incorrect password using LoadOptions. The code catches the expected exception, proving that Aspose.Cells rejects invalid passwords.
class VerifyIncorrectPassword
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Secret data");

        // Set a password to encrypt the workbook
        wb.Settings.Password = "correctPassword";

        // Save the encrypted workbook
        string filePath = "encryptedWorkbook.xlsx";
        wb.Save(filePath);

        // Prepare load options with an incorrect password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "wrongPassword";

        // Try to open the encrypted workbook with the wrong password
        try
        {
            Workbook wbWrong = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened unexpectedly with wrong password.");
        }
        catch (Exception ex)
        {
            // Expected: an exception is thrown because the password is incorrect
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }
    }
}
