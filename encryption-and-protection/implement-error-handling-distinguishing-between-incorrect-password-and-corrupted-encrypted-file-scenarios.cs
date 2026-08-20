// Title: Aspose.Cells .NET – Distinguish Wrong Password from Corrupted Encrypted Excel Workbook
// Description: Shows how to detect an encrypted Excel file, verify the user‑supplied password with FileFormatUtil.VerifyPassword, and load the workbook using LoadOptions. The sample catches a CellsException with ExceptionType.FileCorrupted to signal a damaged file, while a false password result triggers an "Incorrect password" message. All other errors are handled generically.
// Keywords: Aspose.Cells encrypted workbook .NET | verify password Aspose.Cells | FileFormatUtil VerifyPassword example | detect corrupted encrypted Excel file | ExceptionType.FileCorrupted handling | wrong password vs file corruption
// Common Searches: Aspose.Cells differentiate wrong password from corrupted file | catch FileCorrupted exception for encrypted Excel workbook | verify password before opening encrypted workbook Aspose.Cells | handle corrupted encrypted Excel file in C#
// Developer Intent: Identify whether an encrypted Excel file fails to open because the password is incorrect or because the file itself is corrupted.
// Use Cases: Validate a user‑entered password and show a specific "Incorrect password" alert. | Open an encrypted workbook safely and report a "File is corrupted" message when the content is damaged. | Return distinct error codes for password mismatch and file corruption in an automated import pipeline.
// AI Prompts: Create a reusable C# method that uses Aspose.Cells to verify a password, load an encrypted workbook, and return separate status codes for wrong password and corrupted file. | Show how to log detailed diagnostics when a CellsException with ExceptionType.FileCorrupted is thrown while opening an encrypted Excel file. | Provide example code that wraps password verification and workbook loading in try‑catch blocks and displays user‑friendly messages for wrong password, corrupted file, and unexpected errors.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to detect an encrypted Excel file, verify the user‑supplied password with FileFormatUtil.VerifyPassword, and load the workbook using LoadOptions. The sample catches a CellsException with ExceptionType.FileCorrupted to signal a damaged file, while a false password result triggers an "Incorrect password" message. All other errors are handled generically.
class EncryptedFileHandler
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "encrypted.xlsx";

        // Password supplied by the user
        string password = "userPassword";

        // Detect file format and check if the file is encrypted
        FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
        if (!fileInfo.IsEncrypted)
        {
            Console.WriteLine("The file is not encrypted.");
            return;
        }

        // Verify whether the supplied password is correct
        bool isPasswordCorrect;
        using (FileStream stream = File.OpenRead(filePath))
        {
            isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, password);
        }

        if (!isPasswordCorrect)
        {
            // Password does not match the encryption password
            Console.WriteLine("Incorrect password.");
            return;
        }

        // Password is correct – attempt to load the workbook
        try
        {
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
            // Perform further processing with 'workbook' as needed
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
        {
            // The file is encrypted but its content is corrupted
            Console.WriteLine("The encrypted file is corrupted.");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
