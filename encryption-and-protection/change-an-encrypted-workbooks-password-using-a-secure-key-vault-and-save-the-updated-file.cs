// Title: Change an Encrypted Excel Workbook Password with Aspose.Cells and Azure Key Vault (C#)
// Description: Demonstrates how to fetch the current and new passwords from a secure key vault, open an encrypted .xlsx file using LoadOptions.Password, assign a new password via Workbook.Settings.Password, and save the workbook so it is re‑encrypted with the updated credential. Includes basic error handling for missing files and secret retrieval failures.
// Keywords: Aspose.Cells C# password change | encrypted Excel workbook | LoadOptions.Password | Workbook.Settings.Password | Azure Key Vault secret retrieval | programmatic Excel re‑encryption | Excel file password rotation | secure Excel protection | C# Excel encryption update
// Common Searches: How to change the password of an encrypted Excel file using Aspose.Cells C# | Aspose.Cells load workbook with old password and save with new password | Retrieve Excel passwords from Azure Key Vault in .NET | Programmatically re‑encrypt an .xlsx file with a new password | C# example for rotating Excel workbook passwords
// Developer Intent: Replace the existing password of an encrypted Excel workbook with a new password obtained securely from a key vault.
// Use Cases: Automate periodic password rotation to satisfy compliance policies. | Migrate legacy workbooks to a new corporate password standard without manual re‑encryption. | Integrate password updates into CI/CD pipelines for secure Excel artifacts.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an encrypted Excel workbook, fetch the old and new passwords from Azure Key Vault, change the password, and save the file. | Explain the difference between LoadOptions.Password and Workbook.Settings.Password when re‑encrypting an Excel file with Aspose.Cells. | Provide robust error‑handling patterns for missing key‑vault secrets, incorrect old passwords, and file‑access issues in the password‑change workflow.

using System;
using System.IO;
using Aspose.Cells;

namespace ChangeWorkbookPassword
{
    // Demonstrates how to fetch the current and new passwords from a secure key vault, open an encrypted .xlsx file using LoadOptions.Password, assign a new password via Workbook.Settings.Password, and save the workbook so it is re‑encrypted with the updated credential. Includes basic error handling for missing files and secret retrieval failures.
    class Program
    {
        // Placeholder for retrieving secrets from a secure key vault.
        // Replace with actual key vault integration as needed.
        static string GetSecret(string secretName)
        {
            // For demonstration purposes only.
            // In production, fetch the secret securely.
            return secretName switch
            {
                "OldPassword" => "oldPasswordFromVault",
                "NewPassword" => "newPasswordFromVault",
                _ => throw new ArgumentException("Unknown secret name.")
            };
        }

        static void Main()
        {
            // Paths to the source encrypted workbook and the output workbook.
            string sourceFilePath = "encrypted_workbook.xlsx";
            string outputFilePath = "encrypted_workbook_updated.xlsx";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException.
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Retrieve the current (old) password and the desired new password from the key vault.
                string oldPassword = GetSecret("OldPassword");
                string newPassword = GetSecret("NewPassword");

                // Load the encrypted workbook using the old password.
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(sourceFilePath, loadOptions);

                // Change the workbook's encryption password to the new password.
                workbook.Settings.Password = newPassword;

                // Save the workbook; it will be re‑encrypted with the new password.
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
