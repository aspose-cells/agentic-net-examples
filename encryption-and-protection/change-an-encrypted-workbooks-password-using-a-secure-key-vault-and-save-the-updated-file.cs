// Title: Rotate the password of an encrypted XLSX workbook using Aspose.Cells with secrets stored in Azure Key Vault (C#)
// AI Prompts: Load an encrypted Excel workbook using the old password fetched from Azure Key Vault, assign a new password via Workbook.Settings, and save the file with Aspose.Cells in C#. | Implement password rotation for a protected XLSX file by retrieving both old and new passwords from a secret manager and re‑saving the workbook using Aspose.Cells.
// Common Searches: how to programmatically change the password of an encrypted Excel file using Aspose.Cells in C# | load password protected XLSX from Azure Key Vault with Aspose.Cells | replace workbook password and save new encrypted file using Aspose.Cells C# example | rotate Excel file protection password using secret manager and Aspose.Cells | C# code to read old password from Azure Key Vault and set new password for XLSX
// Tags: Aspose.Cells change workbook password C# | load encrypted XLSX with password Aspose.Cells | save XLSX with new encryption password Aspose.Cells | Azure Key Vault secret retrieval for Excel passwords | password rotation for protected Excel files C# | secure Excel file protection using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample retrieves the old and new workbook passwords from a key vault, opens the encrypted XLSX with the old password via Aspose.Cells, updates Workbook.Settings with the new password, and saves the workbook as a newly encrypted file.
class Program
{
    static void Main()
    {
        try
        {
            // Retrieve passwords securely from a key vault (replace with actual implementation)
            string oldPassword = GetSecretFromKeyVault("OldWorkbookPassword");
            string newPassword = GetSecretFromKeyVault("NewWorkbookPassword");

            const string inputPath = "encrypted.xlsx";
            const string outputPath = "encrypted_updated.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the encrypted workbook using the old password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Assign the new password to the workbook settings (this will be used when saving)
            workbook.Settings.Password = newPassword;

            // Prepare save options for XLSX (no password property needed here)
            var saveOptions = new OoxmlSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new password
            try
            {
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Placeholder for secure secret retrieval from a key vault
    static string GetSecretFromKeyVault(string secretName)
    {
        // TODO: Integrate with Azure Key Vault, AWS Secrets Manager, etc.
        // For demonstration, return a dummy value.
        return "YourSecretValue";
    }
}
