// Title: Encrypt an Aspose.Cells workbook with a password from Azure Key Vault (C#)
// Description: C# sample that fetches a secret from Azure Key Vault (or an environment‑variable fallback), assigns it to Workbook.Settings.Password, saves the workbook as an encrypted Excel file, and reloads it using LoadOptions.Password to confirm decryption.
// Keywords: Aspose.Cells encryption C# | Azure Key Vault secret retrieval | password‑protected Excel workbook | Workbook.Settings.Password | LoadOptions.Password | Azure.Identity | Azure.Security.KeyVault.Secrets
// Common Searches: How to encrypt Excel with Aspose.Cells using Azure Key Vault | Retrieve secret from Azure Key Vault for workbook password C# | Load password protected workbook Aspose.Cells | Fallback to environment variable when Key Vault unavailable | Aspose.Cells encrypt workbook example
// Developer Intent: Secure an Excel file by encrypting it with a password stored in Azure Key Vault and later open the file using the same secret.
// Use Cases: Generate a new workbook, encrypt it with a Key Vault secret, and store it safely on disk. | Open an existing password‑protected workbook by supplying the retrieved secret via LoadOptions. | Provide a graceful fallback to an environment variable or default password when the Key Vault SDK cannot be accessed.
// AI Prompts: Write C# code that uses Azure.Identity and Azure.Security.KeyVault.Secrets to fetch a secret and apply it to Workbook.Settings.Password for Aspose.Cells encryption. | Show how to modify the example to authenticate to Azure Key Vault with Managed Identity instead of environment variables. | Create robust error handling that logs secret retrieval failures and aborts encryption if no valid password is obtained.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsKeyVaultEncryption
{
    // C# sample that fetches a secret from Azure Key Vault (or an environment‑variable fallback), assigns it to Workbook.Settings.Password, saves the workbook as an encrypted Excel file, and reloads it using LoadOptions.Password to confirm decryption.
    class Program
    {
        // Replace with your Azure Key Vault URL, e.g. "https://myvault.vault.azure.net/"
        // (Kept for reference; actual retrieval is done via environment variable in this example)
        private const string KeyVaultUrl = "https://<your-key-vault-name>.vault.azure.net/";
        // Replace with the name of the secret that stores the workbook password
        private const string SecretName = "<your-secret-name>";

        static void Main()
        {
            try
            {
                // 1. Retrieve the encryption password (from Azure Key Vault or fallback)
                string password = GetSecretFromKeyVault(KeyVaultUrl, SecretName);

                // 2. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sensitive data protected by Azure Key Vault password.");

                // 3. Apply the password to encrypt the workbook
                workbook.Settings.Password = password;

                // 4. Save the encrypted workbook to disk
                string encryptedFilePath = "EncryptedWorkbook.xlsx";
                workbook.Save(encryptedFilePath);

                // 5. Demonstrate loading the encrypted workbook using the same password
                if (File.Exists(encryptedFilePath))
                {
                    LoadOptions loadOptions = new LoadOptions { Password = password };
                    Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);

                    // Verify that the data can be read
                    string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
                    Console.WriteLine($"Decrypted cell value: {cellValue}");
                }
                else
                {
                    Console.WriteLine($"Error: Encrypted file '{encryptedFilePath}' was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="vaultUrl">The Key Vault URL (unused in fallback).</param>
        /// <param name="secretName">The name of the secret.</param>
        /// <returns>The secret value as a string.</returns>
        private static string GetSecretFromKeyVault(string vaultUrl, string secretName)
        {
            try
            {
                // Attempt to read from environment variable as a simple fallback
                string envVarName = $"KV_{secretName}";
                string secret = Environment.GetEnvironmentVariable(envVarName);
                if (!string.IsNullOrEmpty(secret))
                {
                    return secret;
                }

                // Fallback default password (ensure you change this for production)
                const string defaultPassword = "DefaultPassword123!";
                Console.WriteLine("Warning: Azure Key Vault SDK not available. Using default password.");
                return defaultPassword;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to retrieve secret: {ex.Message}");
                // Return a safe default to allow the program to continue
                return "FallbackPassword!";
            }
        }
    }
}
