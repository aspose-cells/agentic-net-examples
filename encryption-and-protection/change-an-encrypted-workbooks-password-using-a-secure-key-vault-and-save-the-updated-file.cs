// Title: Change an Excel workbook password with Aspose.Cells using a secure key vault (C#)
// Description: Loads an encrypted workbook with the old password fetched from a key vault, replaces the opening password via Workbook.Settings.Password, and saves the file with the new password. Includes robust Aspose.Cells and generic exception handling.
// Keywords: Aspose.Cells change workbook password | C# load encrypted Excel | update Excel opening password | Azure Key Vault secret retrieval | Workbook.Settings.Password | secure Excel password rotation | exception handling Aspose.Cells
// Common Searches: replace password of a protected Excel file using Aspose.Cells C# | load encrypted workbook and set new password programmatically | retrieve Excel passwords from Azure Key Vault in C# | change opening password of Aspose.Cells workbook
// Developer Intent: Replace the current opening password of an encrypted Excel workbook with a new password obtained from a secure key vault and save the updated file.
// Use Cases: Automate periodic password rotation for confidential spreadsheets stored in a document management system. | Re‑encrypt workbooks after CI/CD pipelines update encryption keys to maintain compliance. | Migrate legacy password‑protected Excel files to a centralized secret‑management solution.
// AI Prompts: Write C# code that opens a password‑protected Excel file with Aspose.Cells, reads old and new passwords from Azure Key Vault, updates the workbook password, and saves the file. | Show best‑practice error handling for Aspose.Cells password updates, covering CellsException and generic exceptions. | Replace the placeholder GetSecret method with actual Azure Key Vault SDK calls for retrieving workbook passwords.

using System;
using System.IO;
using Aspose.Cells;

namespace ChangeWorkbookPassword
{
    // Loads an encrypted workbook with the old password fetched from a key vault, replaces the opening password via Workbook.Settings.Password, and saves the file with the new password. Includes robust Aspose.Cells and generic exception handling.
    class Program
    {
        // Placeholder for retrieving secrets from a secure key vault.
        // Replace with actual key vault SDK calls as needed.
        static string GetSecret(string secretName)
        {
            // For demonstration purposes only.
            // In production, fetch the secret securely.
            if (secretName == "OldWorkbookPassword")
                return "oldPassword123";
            if (secretName == "NewWorkbookPassword")
                return "newSecurePassword456";

            throw new ArgumentException($"Secret '{secretName}' not found.");
        }

        static void Main()
        {
            // Paths to the source (encrypted) and destination workbook files.
            string sourceFilePath = "EncryptedWorkbook.xlsx";
            string destinationFilePath = "WorkbookWithNewPassword.xlsx";

            // Verify source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourceFilePath))
            {
                Console.WriteLine($"Source file not found: {sourceFilePath}");
                return;
            }

            // Retrieve passwords from the secure key vault.
            string oldPassword = GetSecret("OldWorkbookPassword");
            string newPassword = GetSecret("NewWorkbookPassword");

            try
            {
                // Load the encrypted workbook using the old password.
                var loadOptions = new LoadOptions
                {
                    Password = oldPassword
                };
                var workbook = new Workbook(sourceFilePath, loadOptions);

                // Change the workbook's opening password to the new one.
                workbook.Settings.Password = newPassword;

                // Save the workbook with the updated password.
                workbook.Save(destinationFilePath);
                Console.WriteLine($"Workbook saved successfully to {destinationFilePath}");
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
