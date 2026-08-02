// Title: Validate that a wrong password cannot open an encrypted Excel file with Aspose.Cells (.NET)
// Description: The example creates a workbook, writes confidential data, encrypts it with a password, saves it, confirms encryption with the correct password, then attempts to load the same file using an incorrect password. The code catches the expected exception and demonstrates how to verify password validity without fully loading the workbook via FileFormatUtil.VerifyPassword.
// Keywords: Aspose.Cells | .NET | Excel encryption | wrong password | exception handling | VerifyPassword | secure workbook loading | password protection | FileFormatUtil | data leakage prevention
// Common Searches: Aspose.Cells open encrypted Excel with wrong password | How to catch exception for invalid password in Aspose.Cells | Verify Excel file password without loading content Aspose | Prevent data exposure when opening protected workbook .NET | Check if Excel workbook password is correct using Aspose.Cells
// Developer Intent: Confirm that loading an encrypted workbook with an invalid password throws an exception and does not reveal any cell values.
// Use Cases: Validate user‑provided passwords before processing sensitive Excel files. | Implement a security layer that blocks access when the password is incorrect. | Quickly test password correctness without the overhead of full workbook parsing.
// AI Prompts: Generate a C# unit test using Aspose.Cells that asserts an exception is thrown when a wrong password opens an encrypted Excel file. | Provide code that logs the specific exception type raised for an invalid password in Aspose.Cells. | Write a C# helper method that returns true if FileFormatUtil.VerifyPassword confirms the password, otherwise false.

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, writes confidential data, encrypts it with a password, saves it, confirms encryption with the correct password, then attempts to load the same file using an incorrect password. The code catches the expected exception and demonstrates how to verify password validity without fully loading the workbook via FileFormatUtil.VerifyPassword.
class TestEncryptedWorkbook
{
    static void Main()
    {
        // Create a new workbook and put some confidential data in a cell
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Secret Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "correctPassword";

        // Save the encrypted workbook to disk
        string filePath = "encrypted.xlsx";
        wb.Save(filePath);

        // Verify that the saved workbook reports as encrypted
        Workbook encryptedWb = new Workbook(filePath, new LoadOptions { Password = "correctPassword" });
        Console.WriteLine("Workbook IsEncrypted: " + encryptedWb.Settings.IsEncrypted);

        // Attempt to open the encrypted workbook with an incorrect password
        try
        {
            LoadOptions wrongOptions = new LoadOptions();
            wrongOptions.Password = "wrongPassword";

            // This should throw an exception because the password is invalid
            Workbook wrongWb = new Workbook(filePath, wrongOptions);

            // If no exception occurs (unlikely), trying to read data would expose it
            Console.WriteLine("Cell value with wrong password: " + wrongWb.Worksheets[0].Cells["A1"].Value);
        }
        catch (Exception ex)
        {
            // Expected path: loading fails and no data is exposed
            Console.WriteLine("Failed to open with incorrect password: " + ex.Message);
        }

        // Alternative check: verify password without loading the workbook content
        using (Stream stream = File.OpenRead(filePath))
        {
            bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, "wrongPassword");
            Console.WriteLine("Password verification (wrong password): " + isPasswordValid);
        }
    }
}
