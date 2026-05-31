using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookPasswordChanger
{
    class ChangeWorkbookPassword
    {
        static void Main()
        {
            try
            {
                // Input encrypted workbook path
                string inputPath = "encrypted.xlsx";
                // Output path for workbook with new password
                string outputPath = "encrypted_new_password.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Retrieve passwords from a secure vault (simulated here)
                string oldPassword = GetSecret("OldWorkbookPassword");
                string newPassword = GetSecret("NewWorkbookPassword");

                // Load the encrypted workbook using the old password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Set the new password for the workbook
                workbook.Settings.Password = newPassword;

                // Save the workbook with the updated password
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with new password to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Simulated method to fetch secrets from a secure key vault
        static string GetSecret(string key)
        {
            // Replace this stub with actual key vault integration (e.g., Azure Key Vault, AWS Secrets Manager)
            if (key == "OldWorkbookPassword") return "oldPass123";
            if (key == "NewWorkbookPassword") return "newPass456";
            return string.Empty;
        }
    }
}