// Title: C# – Verify and Load an Encrypted Excel Workbook with Aspose.Cells
// Description: Shows how to use Aspose.Cells.FileFormatUtil.VerifyPassword to check a password for an encrypted .xlsx file, log the verification result, and conditionally load the workbook with LoadOptions, then read the first worksheet name.
// Keywords: Aspose.Cells | C# verify Excel password | FileFormatUtil VerifyPassword | load encrypted workbook .NET | Excel encryption verification | password validation Aspose | log verification result | LoadOptions password | encrypted .xlsx handling | Aspose.Cells example
// Common Searches: Aspose.Cells verify password C# | How to check Excel file password before loading | FileFormatUtil VerifyPassword usage | Load encrypted Excel workbook with password Aspose | C# code to validate encrypted workbook password | Aspose.Cells password verification example
// Developer Intent: Determine if a supplied password unlocks an encrypted Excel file and open the workbook only when the password is valid.
// Use Cases: Authenticate user‑provided passwords before opening protected spreadsheets in web portals. | Batch decryption of multiple encrypted workbooks in automated ETL pipelines. | Compliance logging of password verification attempts in financial reporting workflows. | Conditional processing of encrypted Excel templates in desktop applications.
// AI Prompts: Generate C# code that uses Aspose.Cells to verify an Excel file password, write the verification outcome to the console, and open the workbook with LoadOptions if the password is correct. | Create a reusable method VerifyAndOpen(string filePath, string password) that returns a Workbook when the password is valid or null otherwise. | Show how to log detailed verification status and then display the name of the first worksheet using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to use Aspose.Cells.FileFormatUtil.VerifyPassword to check a password for an encrypted .xlsx file, log the verification result, and conditionally load the workbook with LoadOptions, then read the first worksheet name.
class VerifyWorkbookEncryptionPassword
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Password to verify
        string passwordToCheck = "test";

        // Verify the password using FileFormatUtil
        bool isPasswordCorrect;
        using (FileStream stream = File.OpenRead(filePath))
        {
            isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, passwordToCheck);
        }

        // Log the verification result
        Console.WriteLine($"Password verification result for '{passwordToCheck}': {isPasswordCorrect}");

        // If the password is correct, load the workbook with the password
        if (isPasswordCorrect)
        {
            LoadOptions loadOptions = new LoadOptions { Password = passwordToCheck };
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully with the provided password.");
            // Example: output the name of the first worksheet
            Console.WriteLine($"First worksheet name: {workbook.Worksheets[0].Name}");
        }
        else
        {
            Console.WriteLine("Failed to verify password. Workbook not loaded.");
        }
    }
}
