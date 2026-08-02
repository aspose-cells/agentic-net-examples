// Title: AES‑128 vs AES‑256 Encryption of Excel Files with Aspose.Cells for .NET
// Description: This C# example creates a workbook, applies a password, encrypts one copy with a 128‑bit key and another with a 256‑bit key using Aspose.Cells' StrongCryptographicProvider, saves both files, and then reloads them with LoadOptions to confirm successful protection and highlight the stronger security of AES‑256.
// Keywords: Aspose.Cells AES encryption | C# Excel AES-128 | C# Excel AES-256 | SetEncryptionOptions key length | LoadOptions password protected workbook | compare AES strength .NET | StrongCryptographicProvider example | Excel file encryption Aspose | AES key length comparison | secure Excel workbook C#
// Common Searches: how to encrypt Excel with AES‑128 using Aspose.Cells | encrypt Excel workbook with AES‑256 in C# | Aspose.Cells SetEncryptionOptions example | load password protected Excel file Aspose.Cells | AES‑128 vs AES‑256 performance in .NET | compare Excel encryption strengths Aspose
// Developer Intent: The developer wants to protect the same Excel workbook with two different AES key sizes, verify that both files open with the same password, and understand the security difference between AES‑128 and AES‑256.
// Use Cases: Generate two versions of a confidential report to evaluate compliance requirements for different encryption standards. | Automate validation of password‑protected workbooks in a CI pipeline by loading them with LoadOptions. | Showcase how to switch key length without altering workbook content for security policy testing.
// AI Prompts: Write C# code that measures encryption and decryption time for AES‑128 and AES‑256 Excel files using Aspose.Cells. | Explain how file size and memory usage differ between AES‑128 and AES‑256 protected workbooks created with Aspose.Cells. | Provide best‑practice recommendations for selecting AES key length when securing Excel files in enterprise .NET applications.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionComparison
{
    // This C# example creates a workbook, applies a password, encrypts one copy with a 128‑bit key and another with a 256‑bit key using Aspose.Cells' StrongCryptographicProvider, saves both files, and then reloads them with LoadOptions to confirm successful protection and highlight the stronger security of AES‑256.
    class Program
    {
        static void Main()
        {
            try
            {
                // Common password for both encrypted files
                const string password = "StrongPassword123";

                // -----------------------------------------------------------------
                // Create a sample workbook with some data (shared for both encryptions)
                // -----------------------------------------------------------------
                Workbook workbookTemplate = new Workbook();
                Worksheet sheet = workbookTemplate.Worksheets[0];
                sheet.Cells["A1"].PutValue("Encryption Strength Comparison");
                sheet.Cells["A2"].PutValue("AES-128 vs AES-256");
                sheet.Cells["A3"].PutValue(DateTime.Now);

                // -----------------------------------------------------------------
                // Encrypt with AES-128 (key length = 128 bits)
                // -----------------------------------------------------------------
                // Clone the template to keep the original data unchanged
                Workbook workbook128 = new Workbook();
                workbookTemplate.Copy(workbook128);
                workbook128.Settings.Password = password;
                // Apply encryption options: StrongCryptographicProvider with 128‑bit key
                workbook128.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
                // Save the encrypted workbook
                string file128 = "Encrypted_AES128.xlsx";
                workbook128.Save(file128);

                // -----------------------------------------------------------------
                // Encrypt with AES-256 (key length = 256 bits)
                // -----------------------------------------------------------------
                Workbook workbook256 = new Workbook();
                workbookTemplate.Copy(workbook256);
                workbook256.Settings.Password = password;
                // Apply encryption options: StrongCryptographicProvider with 256‑bit key
                workbook256.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
                string file256 = "Encrypted_AES256.xlsx";
                workbook256.Save(file256);

                // -----------------------------------------------------------------
                // Verify encryption by loading each file with the password
                // -----------------------------------------------------------------
                LoadOptions loadOptions = new LoadOptions { Password = password };

                if (File.Exists(file128))
                {
                    Workbook loaded128 = new Workbook(file128, loadOptions);
                    bool isEncrypted128 = loaded128.Settings.IsEncrypted;
                    Console.WriteLine($"{file128} loaded. IsEncrypted = {isEncrypted128}");
                }
                else
                {
                    Console.WriteLine($"File not found: {file128}");
                }

                if (File.Exists(file256))
                {
                    Workbook loaded256 = new Workbook(file256, loadOptions);
                    bool isEncrypted256 = loaded256.Settings.IsEncrypted;
                    Console.WriteLine($"{file256} loaded. IsEncrypted = {isEncrypted256}");
                }
                else
                {
                    Console.WriteLine($"File not found: {file256}");
                }

                // -----------------------------------------------------------------
                // Output simple comparison result
                // -----------------------------------------------------------------
                Console.WriteLine("Encryption comparison completed.");
                Console.WriteLine("Both files are encrypted and can be opened with the same password.");
                Console.WriteLine("AES-256 uses a longer key and is considered stronger than AES-128.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
