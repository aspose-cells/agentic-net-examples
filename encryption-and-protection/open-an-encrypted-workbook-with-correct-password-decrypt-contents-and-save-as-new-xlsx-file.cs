// Title: Decrypt a password‑protected Excel workbook using Aspose.Cells for .NET
// Description: Demonstrates how to open an encrypted XLSX file with a known password via LoadOptions, remove the workbook's protection by clearing the Settings.Password property, and save the result as an unencrypted file. Includes file‑existence validation and basic exception handling.
// Keywords: Aspose.Cells decrypt Excel | remove workbook password .NET | load password‑protected XLSX | save unencrypted workbook | C# Excel decryption example | Aspose.Cells LoadOptions password
// Common Searches: open encrypted xlsx with Aspose.Cells C# | how to remove password from Excel file programmatically | Aspose.Cells decrypt workbook example | save protected Excel as plain file using .NET | batch decrypt Excel files Aspose
// Developer Intent: Load a password‑protected workbook, strip its protection, and write it out as a plain XLSX file.
// Use Cases: Automate decryption of incoming Excel reports before data ingestion. | Prepare password‑locked spreadsheets for archival in an unprotected format. | Enable downstream analytics tools to read Excel files without manual password entry.
// AI Prompts: Generate C# code with Aspose.Cells that opens an encrypted .xlsx using a supplied password, removes the protection, and saves a new unprotected file. | Show robust error handling for workbook decryption in a console app using Aspose.Cells. | Create a script that iterates over a folder of password‑protected Excel files, decrypts each with Aspose.Cells, and stores the results in a target directory.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to open an encrypted XLSX file with a known password via LoadOptions, remove the workbook's protection by clearing the Settings.Password property, and save the result as an unencrypted file. Includes file‑existence validation and basic exception handling.
    public class DecryptWorkbookDemo
    {
        public static void Run()
        {
            // Path to the encrypted workbook
            string encryptedFile = "encrypted.xlsx";

            // Desired path for the decrypted workbook
            string decryptedFile = "decrypted.xlsx";

            // Password used to protect the original workbook
            string password = "myPassword";

            // Verify that the encrypted file exists
            if (!File.Exists(encryptedFile))
            {
                Console.WriteLine($"Error: Encrypted file '{encryptedFile}' not found.");
                return;
            }

            try
            {
                // Load the password‑protected workbook
                LoadOptions loadOptions = new LoadOptions { Password = password };
                Workbook workbook = new Workbook(encryptedFile, loadOptions);

                // Remove password protection
                workbook.Settings.Password = null;

                // Save the workbook as a new unencrypted XLSX file
                workbook.Save(decryptedFile);
                Console.WriteLine($"Decrypted workbook saved to '{decryptedFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
