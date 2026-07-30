// Title: Distinguish Wrong Password and Corrupted Encrypted Excel Files with Aspose.Cells .NET
// Description: C# sample that detects an encrypted workbook, validates the supplied password with FileFormatUtil.VerifyPassword, then loads the file using LoadOptions. It catches CellsException for ExceptionType.FileCorrupted to signal a damaged file and ExceptionType.IncorrectPassword for password failures that occur during loading, while also handling unexpected errors.
// Keywords: Aspose.Cells encrypted Excel detection | verify password Aspose.Cells | FileCorrupted exception .NET | IncorrectPassword exception Aspose.Cells | load encrypted workbook C# | Excel file protection handling
// Common Searches: Aspose.Cells differentiate wrong password from corrupted encrypted workbook | C# catch FileCorrupted when opening password‑protected Excel | verify Excel password before loading with Aspose.Cells | handle incorrect password and file corruption in Aspose.Cells .NET
// Developer Intent: Identify encrypted Excel files, confirm the password, and provide distinct error messages for an invalid password versus a corrupted encrypted workbook when loading with Aspose.Cells.
// Use Cases: Use FileFormatUtil.DetectFileFormat to check IsEncrypted before any processing. | Run FileFormatUtil.VerifyPassword on the file stream to pre‑validate the user‑provided password. | Load the workbook with LoadOptions.Password and catch CellsException where ExceptionType.FileCorrupted indicates file damage. | Catch CellsException with ExceptionType.IncorrectPassword for password mismatches that surface during load. | Log unexpected exceptions for diagnostics while presenting user‑friendly messages.
// AI Prompts: Write C# code that opens a password‑protected Excel file using Aspose.Cells, verifies the password, and returns separate messages for an incorrect password and a corrupted file. | Explain how Aspose.Cells maps ExceptionType.IncorrectPassword and ExceptionType.FileCorrupted to specific error scenarios when loading encrypted workbooks. | Provide best‑practice guidelines for logging and user feedback when handling encrypted workbook errors in a production Aspose.Cells application.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that detects an encrypted workbook, validates the supplied password with FileFormatUtil.VerifyPassword, then loads the file using LoadOptions. It catches CellsException for ExceptionType.FileCorrupted to signal a damaged file and ExceptionType.IncorrectPassword for password failures that occur during loading, while also handling unexpected errors.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string filePath = "encrypted.xlsx";

        // Password supplied by the user
        string password = "test";

        // Detect the file format and check if it is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"IsEncrypted: {formatInfo.IsEncrypted}");

        if (!formatInfo.IsEncrypted)
        {
            Console.WriteLine("The file is not encrypted.");
            return;
        }

        // Verify the supplied password without loading the workbook
        bool isPasswordCorrect;
        using (Stream stream = File.OpenRead(filePath))
        {
            isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, password);
        }

        if (!isPasswordCorrect)
        {
            Console.WriteLine("Incorrect password.");
            return;
        }

        // Attempt to load the workbook with the verified password
        try
        {
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
        {
            Console.WriteLine("The encrypted file is corrupted.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
        {
            // This case is rare because we already verified the password,
            // but it handles scenarios where the password passes verification yet fails on load.
            Console.WriteLine("Incorrect password (detected during load).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
