// Title: Automatic Re‑encryption of an Encrypted Excel Workbook Using Aspose.Cells in C#
// Description: Shows how to load a password‑protected .xlsx with Aspose.Cells LoadOptions, change cell data, re‑apply the original password or a stronger algorithm via Workbook.Settings.Password or SetEncryptionOptions, and save the workbook so the modified file stays encrypted.
// Keywords: Aspose.Cells | C# | automatic workbook re‑encryption | load encrypted Excel | Workbook.Settings.Password | SetEncryptionOptions | AES 128 encryption | strong cryptographic provider | batch re‑encrypt Excel files | Excel file security
// Common Searches: re‑encrypt Excel file after editing with Aspose.Cells | load encrypted workbook C# Aspose.Cells | save modified workbook with same password | change encryption algorithm Aspose.Cells | process multiple encrypted Excel files automatically
// Developer Intent: Apply encryption again after modifying a protected workbook to maintain its security.
// Use Cases: Open an encrypted .xlsx, update cells, set Workbook.Settings.Password to the original password, and save. | Upgrade protection by calling Workbook.SetEncryptionOptions with AES‑128 before saving. | Iterate through a folder of encrypted workbooks, apply changes, and re‑encrypt each file in a batch operation.
// AI Prompts: Write C# code that opens a password‑protected Excel workbook with Aspose.Cells, edits several cells, and saves it re‑encrypted using the same password. | Explain how to switch an Aspose.Cells workbook from default encryption to AES‑256 before saving. | Provide a script that scans a directory for encrypted .xlsx files, modifies them, and re‑applies encryption automatically.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSecurityDemo
{
    // Shows how to load a password‑protected .xlsx with Aspose.Cells LoadOptions, change cell data, re‑apply the original password or a stronger algorithm via Workbook.Settings.Password or SetEncryptionOptions, and save the workbook so the modified file stays encrypted.
    public class AutomaticReEncryption
    {
        public static void Run()
        {
            // Path to the source encrypted workbook
            string sourcePath = "EncryptedInput.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // The password used to open the workbook
            string password = "mySecretPwd";

            try
            {
                // Load the workbook with the password (load rule)
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };

                using (Workbook wb = new Workbook(sourcePath, loadOptions))
                {
                    // ----- Perform any modifications -----
                    // Example: write a new value into the first worksheet
                    Worksheet sheet = wb.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Modified at " + DateTime.Now);

                    // ----- Re‑encrypt the workbook -----
                    // Re‑apply the same password (encryption rule)
                    wb.Settings.Password = password;

                    // Optional: set stronger encryption options if desired
                    // wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Save the workbook (save rule)
                    string outputPath = "EncryptedOutput.xlsx";
                    wb.Save(outputPath);

                    Console.WriteLine($"Workbook re‑encrypted and saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for the application
        public static void Main(string[] args)
        {
            try
            {
                AutomaticReEncryption.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
