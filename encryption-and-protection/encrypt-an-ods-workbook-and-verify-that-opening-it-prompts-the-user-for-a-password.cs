// Title: Password‑protect an ODS workbook and enforce prompt on open with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add data, apply AES‑128 password protection, save as ODS, detect encryption, handle a failed open without a password, and load the file with LoadOptions to verify decryption using Aspose.Cells for .NET.
// Keywords: Aspose.Cells ODS encryption | C# password protect ODS | AES 128 ODS Aspose | FileFormatUtil detect encrypted ODS | LoadOptions password ODS | protect spreadsheet .NET | OpenDocument spreadsheet encryption | verify encrypted workbook | Aspose.Cells security features | C# workbook password prompt
// Common Searches: Aspose.Cells how to encrypt ODS file | C# set password for ODS workbook | detect encrypted ODS with Aspose | open password protected ODS using LoadOptions | AES encryption ODS Aspose.Cells | verify workbook IsEncrypted flag | exception when opening protected ODS | sample code password protect ODS .NET
// Developer Intent: Encrypt an ODS workbook with a password and confirm that opening it requires the password.
// Use Cases: Create a workbook, insert sensitive data, set a password, enable AES‑128 encryption, and save as ODS; then use FileFormatUtil to confirm the file is encrypted. | Attempt to instantiate a Workbook from the protected ODS without a password, catch the resulting exception to demonstrate access denial. | Load the encrypted ODS with LoadOptions supplying the correct password, read cell values, and verify Settings.IsEncrypted reflects successful decryption.
// AI Prompts: Generate C# code that encrypts an ODS workbook with a password using Aspose.Cells, applies AES‑128 encryption, saves the file, and programmatically confirms the encryption status. | Write a C# example that tries to open a password‑protected ODS file without a password, handles the exception, then opens it correctly with LoadOptions and reads a cell value. | Explain step‑by‑step how to set strong AES encryption for an ODS file in Aspose.Cells and how to detect that the file requires a password before it can be opened.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Shows how to create a workbook, add data, apply AES‑128 password protection, save as ODS, detect encryption, handle a failed open without a password, and load the file with LoadOptions to verify decryption using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Set password to encrypt the workbook
            string password = "Secret123";
            wb.Settings.Password = password;

            // Optionally set stronger encryption (AES 128)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save as ODS (OpenDocument Spreadsheet)
            string filePath = "EncryptedWorkbook.ods";
            wb.Save(filePath, SaveFormat.ODS);

            // Verify that the file is encrypted
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"Is file encrypted? {info.IsEncrypted}");

            // Attempt to open without password (should fail)
            try
            {
                Workbook withoutPwd = new Workbook(filePath);
                Console.WriteLine("Opened without password (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Opening without password failed as expected: {ex.Message}");
            }

            // Open with the correct password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook wbProtected = new Workbook(filePath, loadOptions);
            Console.WriteLine($"Cell A1 value after decryption: {wbProtected.Worksheets[0].Cells["A1"].StringValue}");

            // Verify workbook reports encryption status after loading
            Console.WriteLine($"Workbook IsEncrypted after load: {wbProtected.Settings.IsEncrypted}");
        }
    }
}
