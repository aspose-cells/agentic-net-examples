// Title: C# – Verify that an Aspose.Cells encrypted workbook rejects an incorrect password
// Description: Creates a workbook, applies a password with strong encryption, saves it, then uses FileFormatUtil.VerifyPassword to confirm the correct password returns true while a similar wrong password returns false. Demonstrates handling the exception when loading with the wrong password and shows successful decryption with the right password.
// Keywords: Aspose.Cells password verification | C# workbook encryption | FileFormatUtil.VerifyPassword example | LoadOptions incorrect password | StrongCryptographicProvider | Excel file protection C# | Aspose.Cells encryption validation
// Common Searches: Aspose.Cells verify workbook password C# | How to check Excel file password with Aspose | Load encrypted workbook with wrong password exception | C# validate Excel encryption using Aspose.Cells | FileFormatUtil.VerifyPassword usage
// Developer Intent: Ensure that a workbook protected with a specific password cannot be opened with any other similar password.
// Use Cases: Pre‑validate a supplied password before opening an encrypted workbook to avoid unnecessary I/O. | Catch and handle the exception thrown when LoadOptions uses an incorrect password, preventing unauthorized access. | Confirm successful decryption by reading cell data after opening with the correct password.
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, encrypts it with a password, and verifies the password using FileFormatUtil.VerifyPassword. | Show how to catch the exception when loading an encrypted workbook with an incorrect password using LoadOptions in Aspose.Cells. | Explain how to confirm that decryption succeeded by reading a cell value after opening the workbook with the correct password.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, applies a password with strong encryption, saves it, then uses FileFormatUtil.VerifyPassword to confirm the correct password returns true while a similar wrong password returns false. Demonstrates handling the exception when loading with the wrong password and shows successful decryption with the right password.
class WorkbookEncryptionValidation
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Encrypted data");

        // Set the encryption password
        string correctPassword = "Secret123";
        wb.Settings.Password = correctPassword;

        // Optionally set stronger encryption options
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        wb.Save(filePath);

        // Verify the password using FileFormatUtil.VerifyPassword (correct password)
        using (FileStream stream = File.OpenRead(filePath))
        {
            bool isCorrect = FileFormatUtil.VerifyPassword(stream, correctPassword);
            Console.WriteLine($"Correct password validation: {isCorrect}");
        }

        // Verify the password using FileFormatUtil.VerifyPassword (similar wrong password)
        using (FileStream stream = File.OpenRead(filePath))
        {
            bool isWrong = FileFormatUtil.VerifyPassword(stream, "Secret124");
            Console.WriteLine($"Wrong password validation: {isWrong}");
        }

        // Attempt to load the workbook with the wrong password and expect failure
        try
        {
            LoadOptions loadOptionsWrong = new LoadOptions();
            loadOptionsWrong.Password = "Secret124";
            Workbook wbWrong = new Workbook(filePath, loadOptionsWrong);
            Console.WriteLine("Loaded with wrong password (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load with wrong password as expected: {ex.Message}");
        }

        // Load the workbook with the correct password to confirm successful decryption
        LoadOptions loadOptionsCorrect = new LoadOptions();
        loadOptionsCorrect.Password = correctPassword;
        Workbook wbCorrect = new Workbook(filePath, loadOptionsCorrect);
        Console.WriteLine($"Loaded with correct password, cell A1 value: {wbCorrect.Worksheets[0].Cells["A1"].StringValue}");
    }
}
