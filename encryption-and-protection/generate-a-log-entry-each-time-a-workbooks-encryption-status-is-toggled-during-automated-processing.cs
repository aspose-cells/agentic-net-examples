// Title: Log Workbook Encryption State Changes with Aspose.Cells for .NET
// Description: A C# sample that creates a workbook, records the initial IsEncrypted flag, enables encryption by assigning a password, logs the updated state, saves and reloads the encrypted file, then disables encryption by clearing the password, logs the pre‑save status, saves and reloads the unencrypted file, and logs the IsEncrypted property after each operation.
// Keywords: Aspose.Cells | C# encryption logging | Workbook IsEncrypted | toggle password Aspose.Cells | track workbook protection | log encryption status .NET | save encrypted workbook | load encrypted workbook
// Common Searches: how to log encryption status in Aspose.Cells | track IsEncrypted when setting password C# | log workbook protection changes Aspose.Cells | monitor workbook encryption toggle .NET | record encryption state during workbook processing
// Developer Intent: Create timestamped log entries whenever a workbook’s encryption is turned on or off during automated workflows.
// Use Cases: Verify that setting a password actually encrypts the workbook before saving. | Confirm that clearing the password removes encryption prior to publishing. | Audit encryption state after loading files to ensure the IsEncrypted flag matches the file content.
// AI Prompts: Generate C# code that writes a timestamped log each time Workbook.Settings.Password is assigned or cleared, displaying the current IsEncrypted value. | Provide a reusable logging helper for Aspose.Cells that records encryption state during workbook lifecycle events (create, save, load). | Show an example that toggles workbook encryption, logs IsEncrypted at every step, and saves both encrypted and unencrypted versions.

using System;
using Aspose.Cells;

namespace EncryptionToggleLogger
{
    // A C# sample that creates a workbook, records the initial IsEncrypted flag, enables encryption by assigning a password, logs the updated state, saves and reloads the encrypted file, then disables encryption by clearing the password, logs the pre‑save status, saves and reloads the unencrypted file, and logs the IsEncrypted property after each operation.
    class Program
    {
        // Simple logger that writes to console with timestamp
        static void Log(string message)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }

        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Log initial encryption status
            Log($"Initial IsEncrypted: {workbook.Settings.IsEncrypted}");

            // 2. Enable encryption by setting a password (toggle on)
            string password = "Secret123";
            workbook.Settings.Password = password;

            // Log status after setting password (before saving)
            Log($"After setting password, IsEncrypted: {workbook.Settings.IsEncrypted}");

            // Save the encrypted workbook (lifecycle: save)
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            Log($"Workbook saved to '{encryptedPath}' with encryption.");

            // Load the encrypted workbook to verify (lifecycle: load)
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook loadedEncrypted = new Workbook(encryptedPath, loadOptions);
            Log($"Loaded encrypted workbook IsEncrypted: {loadedEncrypted.Settings.IsEncrypted}");

            // 3. Disable encryption by clearing the password (toggle off)
            loadedEncrypted.Settings.Password = string.Empty; // remove password

            // Log status after clearing password (before saving)
            Log($"After clearing password, IsEncrypted (pre‑save): {loadedEncrypted.Settings.IsEncrypted}");

            // Save the unencrypted workbook
            string unencryptedPath = "UnencryptedWorkbook.xlsx";
            loadedEncrypted.Save(unencryptedPath, SaveFormat.Xlsx);
            Log($"Workbook saved to '{unencryptedPath}' without encryption.");

            // Load the unencrypted workbook to confirm
            Workbook loadedUnencrypted = new Workbook(unencryptedPath);
            Log($"Loaded unencrypted workbook IsEncrypted: {loadedUnencrypted.Settings.IsEncrypted}");
        }
    }
}
