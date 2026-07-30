// Title: Decrypt a password‑protected Excel workbook with Aspose.Cells for .NET and verify data integrity
// Description: Shows how to validate a workbook password using FileFormatUtil.VerifyPassword, open the encrypted file with LoadOptions.Password, confirm the IsEncrypted flag, read cell A1 to ensure the content is intact, clear the password, and save an unprotected copy.
// Keywords: Aspose.Cells decrypt workbook | C# Excel password verification | FileFormatUtil.VerifyPassword | LoadOptions.Password example | remove Excel encryption .NET | check IsEncrypted property | read cell after decryption
// Common Searches: how to open a password protected xlsx with Aspose.Cells | verify Excel file password before loading in C# | read cell value from encrypted workbook using Aspose | remove password from Excel file programmatically | Aspose.Cells example for decrypting workbook
// Developer Intent: Open a protected Excel file, confirm the password, validate that the data is correct, and save the workbook without encryption.
// Use Cases: Validate user‑supplied passwords before processing confidential spreadsheets. | Extract specific cell values from a secured workbook to confirm content after decryption. | Strip encryption from Excel files for downstream processing or archiving.
// AI Prompts: Write C# code that uses Aspose.Cells to verify a password for an encrypted .xlsx, open the workbook, read cell A1, remove the password, and save the file unencrypted. | Explain the sequence of Aspose.Cells APIs (FileFormatUtil.VerifyPassword, LoadOptions.Password, Workbook.Settings.IsEncrypted, Workbook.Settings.Password) needed to decrypt an Excel workbook and confirm its data integrity.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to validate a workbook password using FileFormatUtil.VerifyPassword, open the encrypted file with LoadOptions.Password, confirm the IsEncrypted flag, read cell A1 to ensure the content is intact, clear the password, and save an unprotected copy.
class DecryptWorkbookDemo
{
    static void Main()
    {
        // Path to the encrypted workbook and its password
        string encryptedFilePath = "encrypted.xlsx";
        string password = "mySecret";

        // Verify that the provided password is correct for the encrypted file
        using (Stream stream = File.OpenRead(encryptedFilePath))
        {
            bool isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, password);
            Console.WriteLine($"Password verification result: {isPasswordCorrect}");
            if (!isPasswordCorrect)
            {
                Console.WriteLine("Incorrect password. Cannot proceed.");
                return;
            }
        }

        // Load the workbook using the correct password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

        // Confirm that the workbook was originally encrypted
        Console.WriteLine($"Workbook IsEncrypted after load: {workbook.Settings.IsEncrypted}");

        // Verify data integrity by reading a known cell value
        string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine($"Cell A1 value after decryption: {cellValue}");

        // Remove the encryption password and save the workbook unprotected
        workbook.Settings.Password = null;
        workbook.Save("decrypted.xlsx");
        Console.WriteLine("Decrypted workbook saved as 'decrypted.xlsx'.");
    }
}
