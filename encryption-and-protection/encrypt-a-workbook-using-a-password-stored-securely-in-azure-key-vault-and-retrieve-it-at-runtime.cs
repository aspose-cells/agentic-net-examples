// Title: C# – Encrypt an Aspose.Cells workbook with Azure Key Vault password using AES‑256 and open it at runtime
// Description: Demonstrates how to fetch a workbook password from Azure Key Vault (simulated with an environment variable), apply AES‑256 encryption via Aspose.Cells, save the file, and reload it with LoadOptions to verify decryption.
// Keywords: Aspose.Cells encrypt workbook C# | AES 256 Excel encryption Aspose | Azure Key Vault secret for Excel password | password‑protected workbook Aspose.Cells | LoadOptions password Aspose.Cells | secure Excel file generation .NET
// Common Searches: encrypt Excel file with Aspose.Cells and Azure Key Vault | Aspose.Cells AES‑256 password protection .NET | load password protected workbook using Aspose.Cells | retrieve secret from Azure Key Vault for Excel encryption | C# example Aspose.Cells workbook encryption
// Developer Intent: Securely encrypt an Excel workbook with a password obtained from Azure Key Vault and later open the same file programmatically using the retrieved secret.
// Use Cases: Generate compliance‑ready reports that are encrypted with a centrally managed secret. | Store sensitive financial data in Excel files protected by Azure Key Vault‑managed passwords. | Automate decryption of password‑protected workbooks for downstream processing without hard‑coding credentials.
// AI Prompts: Show me C# code that uses the Azure Key Vault SDK to read a secret and apply it as the password for Aspose.Cells workbook encryption. | Provide a step‑by‑step guide to encrypt an Aspose.Cells workbook with AES‑256 and then open it using LoadOptions in .NET. | Explain how to handle missing Key Vault secrets gracefully when encrypting an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionWithKeyVault
{
    // Demonstrates how to fetch a workbook password from Azure Key Vault (simulated with an environment variable), apply AES‑256 encryption via Aspose.Cells, save the file, and reload it with LoadOptions to verify decryption.
    class Program
    {
        static void Main()
        {
            try
            {
                // Retrieve the workbook password.
                // In a real scenario this could be fetched from Azure Key Vault.
                // Here we read it from an environment variable for simplicity.
                string workbookPassword = Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD");
                if (string.IsNullOrEmpty(workbookPassword))
                {
                    Console.WriteLine("Environment variable 'WORKBOOK_PASSWORD' not found. Using a default password.");
                    workbookPassword = "Default@123";
                }

                // Create a new workbook and add sample data.
                using (Workbook wb = new Workbook())
                {
                    Worksheet sheet = wb.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Sensitive data protected by Azure Key Vault password.");

                    // Apply encryption password.
                    wb.Settings.Password = workbookPassword;

                    // Set stronger encryption options (AES 256-bit).
                    wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

                    // Save the encrypted workbook.
                    string encryptedFilePath = "EncryptedWorkbook.xlsx";
                    wb.Save(encryptedFilePath);
                    Console.WriteLine($"Encrypted workbook saved to '{encryptedFilePath}'.");
                }

                // Load the encrypted workbook using the password.
                string loadPath = "EncryptedWorkbook.xlsx";
                if (!System.IO.File.Exists(loadPath))
                {
                    Console.WriteLine($"File '{loadPath}' not found. Cannot load workbook.");
                    return;
                }

                LoadOptions loadOptions = new LoadOptions
                {
                    Password = workbookPassword
                };

                using (Workbook loadedWb = new Workbook(loadPath, loadOptions))
                {
                    // Verify that the data can be read.
                    string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
                    Console.WriteLine("Decrypted cell value: " + cellValue);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
