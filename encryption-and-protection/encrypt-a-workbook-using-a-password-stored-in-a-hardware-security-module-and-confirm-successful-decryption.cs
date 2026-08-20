// Title: Encrypt and Decrypt an Aspose.Cells Workbook Using an HSM‑Stored Password (C#)
// Description: Demonstrates how to retrieve a password from a hardware security module, apply AES‑128 encryption to an Aspose.Cells workbook, save it, then reload with LoadOptions, verify the IsEncrypted flag, and read a cell to confirm successful decryption.
// Keywords: Aspose.Cells encrypt workbook | C# Excel encryption | hardware security module password | AES 128 Aspose.Cells | LoadOptions password | IsEncrypted property | secure Excel file C# | HSM integration Aspose.Cells | encrypted workbook verification
// Common Searches: How to encrypt an Excel file with a password from an HSM using Aspose.Cells | C# code to load an encrypted workbook with a password | Set AES‑128 encryption for Aspose.Cells workbook | Check if a loaded workbook is encrypted in Aspose.Cells | Retrieve password from hardware security module for Excel encryption
// Developer Intent: Secure a workbook with an HSM‑derived password, then confirm it can be opened and read correctly.
// Use Cases: Protect confidential data in Excel by encrypting with a password managed by a hardware security module. | Apply strong AES‑128 encryption before distributing or archiving Excel reports. | Programmatically validate that an encrypted workbook can be decrypted and its contents accessed.
// AI Prompts: Write C# code that fetches a password from a hardware security module and uses Aspose.Cells to encrypt a workbook with AES‑128. | Show how to open an Aspose.Cells workbook encrypted with a password, verify the IsEncrypted flag, and read a specific cell value. | Explain the steps to integrate an HSM SDK with Aspose.Cells encryption workflow in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Demonstrates how to retrieve a password from a hardware security module, apply AES‑128 encryption to an Aspose.Cells workbook, save it, then reload with LoadOptions, verify the IsEncrypted flag, and read a cell to confirm successful decryption.
    class Program
    {
        // Placeholder for retrieving the password from a hardware security module (HSM)
        static string GetPasswordFromHSM()
        {
            // In a real scenario, integrate with the HSM SDK/API to fetch the password securely.
            // Here we return a hard‑coded value for demonstration purposes.
            return "HSM_Retrieved_Password123!";
        }

        static void Main(string[] args)
        {
            // Retrieve the encryption password from the HSM
            string hsmPassword = GetPasswordFromHSM();

            // -------------------- Create and encrypt workbook --------------------
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data protected by HSM password.");

            // Apply the password to encrypt the workbook
            wb.Settings.Password = hsmPassword;

            // Optionally set stronger encryption options (AES 128‑bit)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedFilePath = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedFilePath, SaveFormat.Xlsx);

            // -------------------- Load and verify decryption --------------------
            // Prepare load options with the same password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = hsmPassword
            };

            // Load the encrypted workbook using the password
            Workbook wbLoaded = new Workbook(encryptedFilePath, loadOptions);

            // Verify that the workbook is indeed encrypted
            Console.WriteLine($"IsEncrypted after load: {wbLoaded.Settings.IsEncrypted}");

            // Read back the cell value to confirm successful decryption
            string cellValue = wbLoaded.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Decrypted cell value: {cellValue}");
        }
    }
}
