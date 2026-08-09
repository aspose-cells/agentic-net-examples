// Title: Verify an encrypted Excel workbook's password with Aspose.Cells in C# and log the result
// Description: This example opens an Excel file as a stream, uses Aspose.Cells.FileFormatUtil.VerifyPassword to determine if a supplied password unlocks the workbook, writes the boolean outcome to the console, and, when the password is correct, loads the workbook with LoadOptions.Password for further processing.
// Keywords: Aspose.Cells verify password C# | FileFormatUtil VerifyPassword example | check Excel encryption .NET | load encrypted workbook Aspose | Excel password validation without opening file | C# stream password verification Aspose.Cells | Excel file protection check Aspose | Aspose.Cells LoadOptions Password usage | log Excel password verification result | secure Excel processing Aspose .NET
// Common Searches: How to verify password of an encrypted Excel file using Aspose.Cells .NET | Aspose.Cells FileFormatUtil VerifyPassword without loading workbook | Load encrypted Excel workbook after password verification in C# | Log result of Excel password check with Aspose.Cells | C# code to test Excel file password using Aspose
// Developer Intent: Confirm whether a specific password can unlock an encrypted Excel workbook and record the verification outcome.
// Use Cases: Validate a user‑provided password before opening a protected workbook to prevent runtime exceptions. | Batch‑process a folder of encrypted Excel files, logging which files open with a known password. | Integrate password verification into a file‑upload service to reject incorrectly protected Excel documents.
// AI Prompts: Generate C# code that uses Aspose.Cells.FileFormatUtil.VerifyPassword to test a password on an Excel file stream and writes the result to a log file. | Show how to modify the example to try multiple passwords sequentially and load the workbook with the first matching password. | Explain how to handle exceptions that may occur during password verification and workbook loading with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// This example opens an Excel file as a stream, uses Aspose.Cells.FileFormatUtil.VerifyPassword to determine if a supplied password unlocks the workbook, writes the boolean outcome to the console, and, when the password is correct, loads the workbook with LoadOptions.Password for further processing.
class VerifyWorkbookEncryptionPassword
{
    static void Main()
    {
        // Path to the workbook that may be encrypted
        string filePath = "encrypted.xlsx";

        // Password to verify
        string passwordToTest = "test";

        // Verify the password using FileFormatUtil without fully loading the workbook
        bool isPasswordCorrect;
        using (Stream stream = File.OpenRead(filePath))
        {
            isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, passwordToTest);
        }

        // Log the verification result
        Console.WriteLine($"Password verification result for '{passwordToTest}': {isPasswordCorrect}");

        // If the password is correct, optionally load the workbook with the password
        if (isPasswordCorrect)
        {
            LoadOptions loadOptions = new LoadOptions { Password = passwordToTest };
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully with the verified password.");
            // Perform any additional operations on the workbook here
        }
        else
        {
            Console.WriteLine("Failed to verify password. Workbook not loaded.");
        }
    }
}
