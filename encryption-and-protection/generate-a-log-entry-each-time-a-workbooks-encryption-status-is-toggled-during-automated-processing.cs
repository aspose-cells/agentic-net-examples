// Title: Log workbook encryption status changes with Aspose.Cells for .NET
// Description: Demonstrates how to capture the IsEncrypted flag before and after applying or removing a password, using Aspose.Cells to encrypt, decrypt, and log each state transition during automated processing.
// Keywords: Aspose.Cells encryption logging | C# workbook IsEncrypted | track Excel protection changes | Aspose.Cells SetEncryptionOptions | automated Excel security audit
// Common Searches: Aspose.Cells log encryption status .NET | how to detect workbook encryption change in C# | record Excel password toggle with Aspose | track IsEncrypted property during processing | log workbook protection events Aspose.Cells
// Developer Intent: Record every change to a workbook’s encryption state while processing it programmatically.
// Use Cases: Establish a baseline by logging the initial IsEncrypted flag of a newly created workbook. | Log the transition to an encrypted file after setting a password and optional encryption options. | Log the return to an unencrypted state after clearing the password and saving the workbook.
// AI Prompts: Create C# code that writes each IsEncrypted change to a timestamped log file instead of the console using Aspose.Cells. | Show a reusable method that accepts a Workbook, logs its encryption status, and returns the logged message. | Provide an example that wraps password assignment and removal in try‑catch blocks and logs success or failure for each state change.

using System;
using Aspose.Cells;

namespace WorkbookEncryptionLogger
{
    // Demonstrates how to capture the IsEncrypted flag before and after applying or removing a password, using Aspose.Cells to encrypt, decrypt, and log each state transition during automated processing.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (unprotected)
            Workbook wb = new Workbook();

            // Log initial encryption status
            Console.WriteLine($"Initial IsEncrypted: {wb.Settings.IsEncrypted}");

            // Set a password to encrypt the workbook
            wb.Settings.Password = "SecretPassword";

            // Optionally define encryption options (strong encryption, 128-bit key)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            wb.Save(encryptedPath, SaveFormat.Xlsx);

            // Log status after applying password
            Console.WriteLine($"After setting password IsEncrypted: {wb.Settings.IsEncrypted}");

            // Load the encrypted workbook using the password
            LoadOptions loadOpts = new LoadOptions { Password = "SecretPassword" };
            Workbook loadedEncrypted = new Workbook(encryptedPath, loadOpts);

            // Verify that the loaded workbook reports as encrypted
            Console.WriteLine($"Loaded encrypted workbook IsEncrypted: {loadedEncrypted.Settings.IsEncrypted}");

            // Remove encryption by clearing the password
            loadedEncrypted.Settings.Password = null; // or string.Empty
            // Save the workbook without a password
            string decryptedPath = "decrypted.xlsx";
            loadedEncrypted.Save(decryptedPath, SaveFormat.Xlsx);

            // Log status after removing password
            Console.WriteLine($"After removing password IsEncrypted: {loadedEncrypted.Settings.IsEncrypted}");

            // Load the decrypted workbook to confirm encryption is cleared
            Workbook loadedDecrypted = new Workbook(decryptedPath);
            Console.WriteLine($"Loaded decrypted workbook IsEncrypted: {loadedDecrypted.Settings.IsEncrypted}");
        }
    }
}
