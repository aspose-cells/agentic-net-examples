// Title: C# – Distinguish Wrong Password from Corrupted Encrypted Excel Workbook using Aspose.Cells
// Description: Demonstrates how to detect an encrypted Excel file, verify a user‑provided password without loading the workbook, and load it only when the password is correct. The sample catches a CellsException with ExceptionType.FileCorrupted to report a damaged file and provides a generic fallback for other errors.
// Keywords: Aspose.Cells C# | encrypted Excel workbook | verify password Aspose.Cells | FileCorrupted exception | CellsException handling | LoadOptions password | Excel file corruption detection | error handling Aspose.Cells
// Common Searches: Aspose.Cells differentiate wrong password and corrupted file | C# verify encrypted Excel password before opening | catch CellsException FileCorrupted Aspose.Cells | how to detect corrupted encrypted workbook with Aspose | load encrypted xlsx with password validation Aspose
// Developer Intent: Add robust error handling that separately reports an incorrect password and a corrupted encrypted workbook when opening an Excel file with Aspose.Cells.
// Use Cases: Validate a password entered by a user without loading the workbook and show a clear "Incorrect password" message. | Identify a damaged encrypted file by catching CellsException with ExceptionType.FileCorrupted and inform the user. | Provide a generic catch‑all for unexpected exceptions during encrypted file processing.
// AI Prompts: Generate C# code that opens an encrypted Excel file with Aspose.Cells, verifies the password first, and returns distinct messages for wrong password and corrupted file. | Show how to catch CellsException with ExceptionType.FileCorrupted when loading an encrypted workbook and log detailed error information. | Write a method that returns status codes for password failure, file corruption, and successful load using Aspose.Cells in C#.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to detect an encrypted Excel file, verify a user‑provided password without loading the workbook, and load it only when the password is correct. The sample catches a CellsException with ExceptionType.FileCorrupted to report a damaged file and provides a generic fallback for other errors.
public class EncryptedFileHandler
{
    public static void Run()
    {
        // Path to the encrypted workbook and the password to test
        string filePath = "encrypted.xlsx";
        string password = "userProvidedPassword";

        // Ensure the file exists before proceeding
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Detect file format and check if the file is encrypted
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
            if (!fileInfo.IsEncrypted)
            {
                Console.WriteLine("The file is not encrypted.");
                return;
            }

            // Verify the password without loading the workbook
            bool passwordIsCorrect;
            using (Stream stream = File.OpenRead(filePath))
            {
                passwordIsCorrect = FileFormatUtil.VerifyPassword(stream, password);
            }

            if (!passwordIsCorrect)
            {
                Console.WriteLine("Incorrect password.");
                return;
            }

            // Password is correct – attempt to load the workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
            // Further processing can be done here
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
        {
            // Corrupted encrypted file scenario
            Console.WriteLine("The encrypted file appears to be corrupted.");
        }
        catch (Exception ex)
        {
            // Any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            EncryptedFileHandler.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
