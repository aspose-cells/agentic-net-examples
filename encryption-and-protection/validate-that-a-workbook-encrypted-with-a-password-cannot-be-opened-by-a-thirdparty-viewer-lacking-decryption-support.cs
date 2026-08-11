// Title: Validate that a password‑protected Aspose.Cells workbook cannot be opened without the correct password
// Description: Creates a workbook, writes sensitive data, encrypts it with a password, saves the file, uses FileFormatUtil to confirm the encrypted flag, attempts to load the file without a password (expecting an exception), then opens it with LoadOptions using the correct password and reads the original cell value.
// Keywords: Aspose.Cells password encryption | detect encrypted Excel file | open encrypted workbook exception | LoadOptions password Aspose.Cells | verify workbook IsEncrypted flag | C# Aspose.Cells encryption example
// Common Searches: how to check if an Excel file is encrypted with Aspose.Cells | exception when opening password‑protected workbook without password Aspose.Cells .NET | load password‑protected workbook using LoadOptions Aspose.Cells | Aspose.Cells detect encrypted workbook before loading
// Developer Intent: Ensure that a workbook protected by a password throws an error when accessed without the password and opens successfully when the correct password is supplied.
// Use Cases: Programmatically confirm the encryption status of a saved workbook. | Validate exception handling for unauthorized access to a password‑protected file. | Read data from an encrypted workbook after providing the correct password.
// AI Prompts: Show me C# code that verifies an Aspose.Cells workbook is encrypted and throws an exception when opened without a password. | How can I open a password‑protected Excel file with Aspose.Cells using LoadOptions and retrieve a cell value?

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionValidation
{
    // Creates a workbook, writes sensitive data, encrypts it with a password, saves the file, uses FileFormatUtil to confirm the encrypted flag, attempts to load the file without a password (expecting an exception), then opens it with LoadOptions using the correct password and reads the original cell value.
    class Program
    {
        static void Main()
        {
            // Path for the encrypted workbook
            string encryptedPath = "encrypted_workbook.xlsx";

            // 1. Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // 2. Set a password to encrypt the workbook
            wb.Settings.Password = "Secret123";

            // 3. Save the workbook (it will be encrypted)
            wb.Save(encryptedPath);
            Console.WriteLine($"Workbook saved and encrypted at: {encryptedPath}");

            // 4. Verify that the file is reported as encrypted using FileFormatInfo
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedPath);
            Console.WriteLine($"FileFormatInfo.IsEncrypted: {formatInfo.IsEncrypted}");

            // 5. Attempt to open the encrypted workbook without providing a password
            try
            {
                // This should throw an exception because the password is missing
                Workbook wbNoPassword = new Workbook(encryptedPath);
                Console.WriteLine("Unexpectedly opened encrypted workbook without password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open without password (expected): {ex.Message}");
            }

            // 6. Open the encrypted workbook with the correct password using LoadOptions
            LoadOptions loadOptions = new LoadOptions { Password = "Secret123" };
            Workbook wbWithPassword = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine($"Opened with password. Settings.IsEncrypted: {wbWithPassword.Settings.IsEncrypted}");

            // 7. Verify that the data is accessible
            string cellValue = wbWithPassword.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Cell A1 value after decryption: {cellValue}");
        }
    }
}
