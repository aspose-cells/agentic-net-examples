// Title: Verify that a password‑protected Excel workbook saved with Aspose.Cells for .NET opens on another machine
// Description: C# example that creates a workbook, writes data, applies a password and strong encryption, saves it as XLSX, then reloads the file with LoadOptions on a simulated remote machine, checks the IsEncrypted flag, and reads the cell to confirm successful decryption.
// Keywords: Aspose.Cells encrypt workbook C# | password protected Excel .NET | load encrypted workbook LoadOptions | verify workbook encryption Aspose | IsEncrypted property Aspose.Cells | strong cryptographic provider 128‑bit
// Common Searches: Aspose.Cells open password protected Excel on another computer | C# verify encrypted workbook can be opened with same password | LoadOptions password Excel file Aspose.Cells | check IsEncrypted after saving workbook | strong encryption Aspose.Cells .NET example
// Developer Intent: Ensure that a workbook saved with a password can be opened and read on a different system using the same password.
// Use Cases: Create and distribute a password‑protected Excel file while guaranteeing recipients can open it. | Automated testing of encryption settings in CI pipelines across multiple machines. | Validate that the IsEncrypted flag persists after saving and reloading the workbook.
// AI Prompts: Provide C# code that encrypts an Excel workbook with a 128‑bit strong cryptographic provider using Aspose.Cells, saves it, and then confirms it can be opened with LoadOptions password. | Generate a step‑by‑step tutorial for testing that a password‑protected workbook saved on one machine opens on another machine with Aspose.Cells for .NET. | Show how to programmatically copy an encrypted workbook to a different location and verify the password works when loading it.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionVerification
{
    // C# example that creates a workbook, writes data, applies a password and strong encryption, saves it as XLSX, then reloads the file with LoadOptions on a simulated remote machine, checks the IsEncrypted flag, and reads the cell to confirm successful decryption.
    class Program
    {
        static void Main()
        {
            // Path for the encrypted workbook
            string encryptedFilePath = "EncryptedWorkbook.xlsx";
            // Password to protect the workbook
            string password = "SecurePass123";

            // ---------- Create and encrypt the workbook ----------
            // Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted content");

            // Set the password for opening the workbook
            wb.Settings.Password = password;

            // (Optional) Set stronger encryption options
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            wb.Save(encryptedFilePath, SaveFormat.Xlsx);

            // Verify that the workbook is marked as encrypted
            Console.WriteLine($"After saving, IsEncrypted: {wb.Settings.IsEncrypted}");

            // ---------- Load the workbook on another machine ----------
            // Simulate loading on a different machine by using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Load the encrypted workbook
            Workbook loadedWb = new Workbook(encryptedFilePath, loadOptions);

            // Check that the workbook reports being encrypted
            Console.WriteLine($"Loaded workbook IsEncrypted: {loadedWb.Settings.IsEncrypted}");

            // Verify that the data can be read correctly
            string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Cell A1 value after decryption: {cellValue}");
        }
    }
}
