// Title: Detect encrypted Excel workbook with Aspose.Cells for .NET and show a warning
// Description: C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to check the IsEncrypted flag of an Excel file. If the workbook is password‑protected, a warning is printed; otherwise the file is loaded normally and the encryption status is confirmed via Workbook.Settings.IsEncrypted.
// Keywords: Aspose.Cells | C# | detect encrypted workbook | FileFormatUtil IsEncrypted | password protected Excel | load encrypted workbook | Excel encryption detection .NET | Workbook.Settings.IsEncrypted | Excel file security
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells | Aspose.Cells detect password protected workbook before loading | C# code to display warning for encrypted Excel workbook | FileFormatUtil DetectFileFormat encryption flag example | load encrypted Excel file with password Aspose.Cells
// Developer Intent: Identify whether an Excel workbook is encrypted before opening it and alert the user if a password is required.
// Use Cases: Validate encryption of user‑uploaded Excel files prior to data extraction. | Prompt for a password only when the workbook is password‑protected. | Log encryption detection events for compliance and audit trails.
// AI Prompts: Generate C# code that checks an Excel file for encryption with Aspose.Cells and returns a boolean result. | Write a method that loads an encrypted workbook using a supplied password after detecting encryption. | Create error‑handling logic for prompting the user when Aspose.Cells detects a password‑protected Excel file.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to check the IsEncrypted flag of an Excel file. If the workbook is password‑protected, a warning is printed; otherwise the file is loaded normally and the encryption status is confirmed via Workbook.Settings.IsEncrypted.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be examined
            string filePath = "sample.xlsx";

            // Detect the file format and check if it is encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            if (formatInfo.IsEncrypted)
            {
                // Encryption detected – display a warning
                Console.WriteLine("Warning: The workbook is encrypted and requires a password to open.");
                // If you have the password, you can load the workbook like this:
                // LoadOptions loadOptions = new LoadOptions { Password = "yourPassword" };
                // Workbook encryptedWorkbook = new Workbook(filePath, loadOptions);
            }
            else
            {
                // No encryption – load the workbook normally
                Workbook workbook = new Workbook(filePath);
                Console.WriteLine("Workbook loaded successfully. Encryption status: " + workbook.Settings.IsEncrypted);
                // Continue processing the workbook as needed
            }
        }
    }
}
