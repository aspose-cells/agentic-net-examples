// Title: Encrypt an Aspose.Cells workbook with an HSM‑derived password and verify decryption (C#)
// Description: Demonstrates retrieving a password from a hardware security module, applying AES‑128 encryption to a new Workbook via SetEncryptionOptions, saving it, then loading it with LoadOptions to confirm the IsEncrypted flag and read back the data.
// Keywords: Aspose.Cells | C# workbook encryption | hardware security module | HSM password | AES 128 encryption | SetEncryptionOptions | LoadOptions | encrypted Excel file .NET | password‑protected workbook | secure Excel storage
// Common Searches: Aspose.Cells encrypt workbook with HSM password C# | How to use SetEncryptionOptions for AES‑128 in Aspose.Cells | Load encrypted Excel file using LoadOptions Aspose.Cells .NET | Verify IsEncrypted flag after opening a protected workbook | Retrieve encryption key from hardware security module for Excel
// Developer Intent: Secure a workbook with an HSM‑provided password and ensure it can be opened programmatically.
// Use Cases: Protect confidential data in Excel by encrypting the file with a key managed in an HSM before distribution. | Automate compliance‑driven storage of financial reports using AES‑128 encryption via Aspose.Cells. | Validate that only applications possessing the HSM‑derived password can decrypt and read the workbook.
// AI Prompts: Write C# code that fetches a password from a hardware security module and encrypts an Aspose.Cells workbook with AES‑256. | Show how to catch and log exceptions when opening an Aspose.Cells workbook encrypted with an HSM password. | Provide a step‑by‑step tutorial for confirming workbook encryption status and reading cell values after decryption in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsEncryptionDemo
{
    // Demonstrates retrieving a password from a hardware security module, applying AES‑128 encryption to a new Workbook via SetEncryptionOptions, saving it, then loading it with LoadOptions to confirm the IsEncrypted flag and read back the data.
    class Program
    {
        // Simulated method that retrieves the password from a hardware security module (HSM)
        // In a real scenario this would interface with the HSM SDK/API.
        static string GetPasswordFromHSM()
        {
            // Placeholder: replace with actual HSM call
            return "HSM_Retrieved_Password123!";
        }

        static void Main(string[] args)
        {
            // Retrieve encryption password from HSM
            string password = GetPasswordFromHSM();

            // ------------------- Create and encrypt workbook -------------------
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data encrypted with HSM password.");

            // Set the workbook password (encryption)
            workbook.Settings.Password = password;

            // Optional: specify stronger encryption options (AES 128-bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedFilePath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

            // ------------------- Load and verify decryption -------------------
            // Prepare load options with the same password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;

            // Load the encrypted workbook
            Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);

            // Verify that the workbook is indeed encrypted
            Console.WriteLine($"IsEncrypted: {loadedWorkbook.Settings.IsEncrypted}");

            // Verify that the data can be read correctly
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Decrypted cell value: {cellValue}");

            // Clean up
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}
