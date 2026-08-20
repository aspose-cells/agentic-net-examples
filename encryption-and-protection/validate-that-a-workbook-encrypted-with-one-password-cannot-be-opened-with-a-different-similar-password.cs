// Title: Validate that an Aspose.Cells workbook encrypted with a password cannot be opened with another password (C#)
// Description: C# sample that creates an Excel workbook, sets a password via Workbook.Settings.Password, saves it, verifies the IsEncrypted flag, uses FileFormatUtil.VerifyPassword to confirm the correct password succeeds and a similar wrong password fails, attempts to load the file with the wrong password (expecting an exception), and finally loads it with the correct password to read a cell value.
// Keywords: Aspose.Cells | .NET | C# | Workbook encryption | Password protection | FileFormatUtil.VerifyPassword | LoadOptions.Password | Excel encryption validation | IsEncrypted flag | incorrect password handling
// Common Searches: How to verify an encrypted Excel workbook password with Aspose.Cells | Aspose.Cells verify password example C# | Exception thrown when loading a password‑protected workbook with wrong password | Check if workbook is encrypted after saving using Aspose.Cells | Validate Excel file password using Aspose.Cells .NET
// Developer Intent: Confirm that a workbook protected with a specific password rejects any other password.
// Use Cases: Programmatically validate passwords with FileFormatUtil.VerifyPassword – true for the correct password, false for a similar incorrect one. | Attempt to open an encrypted workbook using LoadOptions.Password set to a wrong value and handle the resulting exception. | Load the workbook with the correct password to ensure successful decryption and access to cell data.
// AI Prompts: Provide C# code that demonstrates how to verify that an Aspose.Cells workbook encrypted with one password cannot be opened with another password. | Explain the interaction between FileFormatUtil.VerifyPassword and LoadOptions.Password for workbook encryption validation in Aspose.Cells. | Show how to catch and handle the exception thrown when loading a password‑protected workbook with an incorrect password using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that creates an Excel workbook, sets a password via Workbook.Settings.Password, saves it, verifies the IsEncrypted flag, uses FileFormatUtil.VerifyPassword to confirm the correct password succeeds and a similar wrong password fails, attempts to load the file with the wrong password (expecting an exception), and finally loads it with the correct password to read a cell value.
class WorkbookPasswordValidationDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

        // Set the encryption password
        string correctPassword = "Secret123";
        wb.Settings.Password = correctPassword;

        // Save the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        wb.Save(filePath);

        // Verify that the workbook is marked as encrypted
        Console.WriteLine("Workbook IsEncrypted after save: " + wb.Settings.IsEncrypted);

        // Use FileFormatUtil.VerifyPassword to check the correct password
        using (FileStream stream = File.OpenRead(filePath))
        {
            bool isValid = FileFormatUtil.VerifyPassword(stream, correctPassword);
            Console.WriteLine($"Password '{correctPassword}' validation result: {isValid}");
        }

        // Use FileFormatUtil.VerifyPassword to check an incorrect, similar password
        using (FileStream stream = File.OpenRead(filePath))
        {
            string wrongPassword = "Secret124";
            bool isValid = FileFormatUtil.VerifyPassword(stream, wrongPassword);
            Console.WriteLine($"Password '{wrongPassword}' validation result: {isValid}");
        }

        // Attempt to load the workbook with the wrong password (should fail)
        try
        {
            LoadOptions wrongLoad = new LoadOptions();
            wrongLoad.Password = "Secret124";
            Workbook wbWrong = new Workbook(filePath, wrongLoad);
            Console.WriteLine("Loaded with wrong password (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to load with wrong password as expected: " + ex.Message);
        }

        // Load the workbook with the correct password
        LoadOptions correctLoad = new LoadOptions();
        correctLoad.Password = correctPassword;
        Workbook wbLoaded = new Workbook(filePath, correctLoad);
        Console.WriteLine("Loaded with correct password successfully. Cell A1 value: " + wbLoaded.Worksheets[0].Cells["A1"].StringValue);
    }
}
