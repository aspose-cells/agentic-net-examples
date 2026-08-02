// Title: Decrypt a password‑protected Excel workbook and save as XLSX using Aspose.Cells for .NET
// Description: Load an encrypted .xlsx with the correct password via Aspose.Cells LoadOptions, clear the opening password (workbook.Settings.Password = null), and save the workbook as an unprotected XLSX file. Includes robust error handling for CellsException and generic exceptions.
// Keywords: Aspose.Cells decrypt Excel | open password protected workbook C# | remove Excel file password Aspose | load encrypted XLSX .NET | save unprotected workbook Aspose.Cells | Excel encryption removal C# | LoadOptions password Aspose.Cells | Workbook.Settings.Password null
// Common Searches: How to open a password‑protected Excel file with Aspose.Cells | Remove opening password from .xlsx using C# | Decrypt an encrypted workbook and save as new file Aspose | Aspose.Cells load encrypted workbook example | C# code to clear Excel file password
// Developer Intent: Open a password‑protected Excel workbook, strip its opening password, and write the result as an unencrypted XLSX file.
// Use Cases: Automate bulk decryption of incoming encrypted reports before data processing. | Integrate a decryption step in a CI/CD pipeline that generates Excel outputs without passwords. | Validate user‑provided passwords for protected workbooks prior to performing analytics.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an encrypted .xlsx with a known password, remove the password, and save it as a new file. | Explain best practices for handling incorrect passwords and other exceptions when loading protected workbooks with Aspose.Cells. | Create a reusable method: DecryptWorkbook(string encryptedPath, string password, string outputPath) → bool, using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Load an encrypted .xlsx with the correct password via Aspose.Cells LoadOptions, clear the opening password (workbook.Settings.Password = null), and save the workbook as an unprotected XLSX file. Includes robust error handling for CellsException and generic exceptions.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook
            string encryptedFilePath = "encrypted.xlsx";

            // Verify that the source file exists
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"Error: File \"{encryptedFilePath}\" not found.");
                return;
            }

            // Password used to protect the workbook
            string password = "mySecretPassword";

            try
            {
                // Load the encrypted workbook using LoadOptions with the correct password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                // Remove opening password protection
                workbook.Settings.Password = null;

                // Save the decrypted workbook as a new XLSX file
                string decryptedFilePath = "decrypted.xlsx";
                workbook.Save(decryptedFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Decrypted workbook saved to \"{decryptedFilePath}\".");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
