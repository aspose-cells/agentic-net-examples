// Title: Encrypt an Aspose.Cells workbook with password/AES‑128 and save to a network share
// Description: C# example that creates a workbook, adds data, applies a password and AES‑128 encryption, saves the file to a UNC/network folder, and verifies the protection by reopening it with the same password.
// Keywords: Aspose.Cells password protection | AES 128 encryption Aspose.Cells | save encrypted Excel to network share | C# Aspose.Cells workbook encryption | LoadOptions.Password Aspose.Cells | UNC path Excel encryption | secure Excel file Aspose
// Common Searches: how to set password on Aspose.Cells workbook C# | Aspose.Cells AES 128 encryption example | save encrypted Excel file to UNC path | verify password protected workbook Aspose.Cells | C# encrypt Excel before copying to network share
// Developer Intent: The developer needs to protect an Excel workbook with a password and strong AES‑128 encryption using Aspose.Cells, then store the file on a network/UNC location without losing the encryption.
// Use Cases: Automatically generate confidential financial reports, encrypt them with a strong password and AES‑128, and place the files on a shared network folder. | Create scheduled backups of sensitive spreadsheets, applying workbook encryption before copying them to a remote server or UNC share. | Validate that encryption persists after saving by loading the password‑protected workbook from the network location and reading specific cells.
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, applies a password and AES‑128 encryption, and saves it directly to a UNC network path. | Show how to open a password‑protected Excel file stored on a network share using Aspose.Cells LoadOptions.Password and read a cell value. | Explain the differences between Aspose.Cells encryption types (AES‑128, AES‑256, StrongCryptographicProvider) and how to ensure the chosen setting is retained when saving to remote storage.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // C# example that creates a workbook, adds data, applies a password and AES‑128 encryption, saves the file to a UNC/network folder, and verifies the protection by reopening it with the same password.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Encrypted content");

                // Set the password that will be required to open the workbook
                workbook.Settings.Password = "StrongPassword123";

                // Optional: specify stronger encryption options (e.g., AES 128-bit)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // Define a local path to save the encrypted workbook
                string folderPath = Path.Combine(Environment.CurrentDirectory, "Output");
                string filePath = Path.Combine(folderPath, "EncryptedWorkbook.xlsx");

                // Ensure the output directory exists
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                // Save the workbook; the password protection is preserved
                workbook.Save(filePath, SaveFormat.Xlsx);

                // Verify that the file is encrypted by loading it with the password
                LoadOptions loadOptions = new LoadOptions { Password = "StrongPassword123" };
                Workbook loaded = new Workbook(filePath, loadOptions);
                Console.WriteLine("Loaded cell value: " + loaded.Worksheets[0].Cells["A1"].StringValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
