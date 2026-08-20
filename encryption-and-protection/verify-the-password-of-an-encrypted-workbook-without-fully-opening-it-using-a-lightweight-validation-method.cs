// Title: Check an Encrypted Excel Workbook Password in C# Using Aspose.Cells Without Loading the File
// Description: Shows how to open an encrypted .xlsx as a read‑only stream and call Aspose.Cells.FileFormatUtil.VerifyPassword to confirm whether a supplied password is correct, eliminating the need to fully parse the workbook.
// Keywords: Aspose.Cells VerifyPassword | C# Excel password check | encrypted workbook validation | FileFormatUtil VerifyPassword stream | lightweight Excel password verification | Excel file encryption check .NET | password validation without opening workbook
// Common Searches: how to test password of encrypted Excel file using Aspose.Cells | verify Excel workbook password from stream C# | Aspose.Cells password validation without loading workbook | check if .xlsx is protected with password in .NET | quick password check for encrypted Excel file
// Developer Intent: Determine whether a specific password can unlock an encrypted Excel workbook without fully loading the document.
// Use Cases: Validate a user‑entered password before attempting to open a protected spreadsheet. | Scan a collection of encrypted files and flag those that match a known password. | Perform a server‑side pre‑check of a template’s password before executing further processing.
// AI Prompts: Generate C# code that uses Aspose.Cells FileFormatUtil.VerifyPassword to test a password from a file stream and returns a boolean. | Provide an example of handling an incorrect password result from VerifyPassword, including logging and user notification. | Create a reusable method that accepts a file path and password, opens the file as a read‑only stream, and returns true if the password is valid using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    // Shows how to open an encrypted .xlsx as a read‑only stream and call Aspose.Cells.FileFormatUtil.VerifyPassword to confirm whether a supplied password is correct, eliminating the need to fully parse the workbook.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook
            string filePath = "encrypted.xlsx";

            // Password to validate
            string passwordToTest = "test";

            // Open the file as a read‑only stream (no full workbook loading)
            using (Stream stream = File.OpenRead(filePath))
            {
                // Verify the password using the lightweight method
                bool isValid = FileFormatUtil.VerifyPassword(stream, passwordToTest);

                Console.WriteLine($"Password '{passwordToTest}' is valid: {isValid}");
            }
        }
    }
}
