// Title: C# – Verify Excel Workbook Password with Aspose.Cells FileFormatUtil
// Description: Shows how to detect if an .xlsx file is encrypted and validate a supplied password using Aspose.Cells for .NET, then write the results to the console.
// Keywords: Aspose.Cells | .NET | C# | Excel encryption detection | password validation | FileFormatUtil | VerifyPassword method | encrypted workbook | audit logging | batch processing
// Common Searches: Aspose.Cells verify Excel file password C# | detect encrypted .xlsx using Aspose.Cells | FileFormatUtil VerifyPassword example | check if Excel workbook is password protected .NET | log encryption status Aspose.Cells console app
// Developer Intent: Determine whether an Excel file is password‑protected and confirm that a given password is correct.
// Use Cases: Prompt users for a password and open the workbook only when the password matches. | Run a scheduled scan of a folder of spreadsheets, flag files with incorrect or missing passwords, and record the findings. | Produce compliance reports that list encryption status and password verification outcomes for each processed file.
// AI Prompts: Write a C# console program that uses Aspose.Cells to detect encryption on an .xlsx file and verify a user‑provided password, handling both valid and invalid cases. | Create a code snippet that integrates Aspose.Cells password verification into an ASP.NET file‑upload endpoint, returning an error response when the password is incorrect. | Explain how to log the results of FileFormatUtil.VerifyPassword and FileFormatUtil.DetectFileFormat to a structured audit file in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to detect if an .xlsx file is encrypted and validate a supplied password using Aspose.Cells for .NET, then write the results to the console.
class VerifyWorkbookPassword
{
    static void Main()
    {
        // Path to the workbook that may be encrypted
        string filePath = "encrypted.xlsx";

        // Password to verify
        string password = "test";

        // Detect whether the file is encrypted
        using (FileStream detectStream = File.OpenRead(filePath))
        {
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(detectStream);
            Console.WriteLine($"File encrypted: {formatInfo.IsEncrypted}");
        }

        // Verify the provided password against the encrypted workbook
        using (FileStream verifyStream = File.OpenRead(filePath))
        {
            bool isPasswordCorrect = FileFormatUtil.VerifyPassword(verifyStream, password);
            Console.WriteLine($"Password verification result for '{password}': {isPasswordCorrect}");
        }
    }
}
