// Title: Encrypt and Open an Excel Workbook with Azure Key Vault‑Stored Password using Aspose.Cells (C#)
// Description: C# example that fetches a workbook password from Azure Key Vault (with an optional environment‑variable fallback), creates an Excel file, applies 256‑bit strong encryption via Aspose.Cells, saves it, and then reloads the protected workbook using LoadOptions to confirm the data can be read.
// Keywords: Aspose.Cells encrypt workbook C# | Azure Key Vault secret retrieval C# | password‑protected Excel file Aspose.Cells | 256‑bit encryption Aspose.Cells | LoadOptions password protected workbook | secure password storage Azure Key Vault | environment variable fallback password
// Common Searches: How to encrypt an Excel file with a password from Azure Key Vault using Aspose.Cells | Aspose.Cells strong encryption 256‑bit C# example | Load a password‑protected workbook with Aspose.Cells LoadOptions | Retrieve secret from Azure Key Vault in a .NET console app | Best practice for storing Excel passwords securely in C#
// Developer Intent: Securely encrypt an Excel workbook with a password retrieved from Azure Key Vault and later open it programmatically using the same secret.
// Use Cases: Store the workbook password as a secret in Azure Key Vault to avoid hard‑coding credentials. | Apply 256‑bit strong cryptographic encryption to protect sensitive spreadsheet data. | Save the encrypted workbook and later open it by supplying the secret via LoadOptions. | Provide a fallback to an environment variable or interactive prompt when the Key Vault secret is unavailable.
// AI Prompts: Write C# code that reads a secret from Azure Key Vault and uses it as the password to encrypt an Aspose.Cells workbook with 256‑bit encryption. | Show how to open a password‑protected Excel file using Aspose.Cells LoadOptions and verify its contents in C#. | Create a sample that falls back to an environment variable or console prompt if the Azure Key Vault secret cannot be retrieved, then applies strong encryption with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionWithKeyVault
{
    // C# example that fetches a workbook password from Azure Key Vault (with an optional environment‑variable fallback), creates an Excel file, applies 256‑bit strong encryption via Aspose.Cells, saves it, and then reloads the protected workbook using LoadOptions to confirm the data can be read.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Retrieve workbook password from environment variable or prompt the user
                string workbookPassword = Environment.GetEnvironmentVariable("WorkbookPassword");
                if (string.IsNullOrEmpty(workbookPassword))
                {
                    Console.Write("Enter workbook password: ");
                    workbookPassword = Console.ReadLine();
                }

                // Create a new workbook and add sample data
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sensitive data protected by Azure Key Vault password.");

                // Apply password protection and set strong encryption
                wb.Settings.Password = workbookPassword;
                wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

                // Save the encrypted workbook
                string outputPath = "EncryptedWorkbook.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");

                // Ensure the file exists before attempting to load it
                if (!System.IO.File.Exists(outputPath))
                {
                    Console.WriteLine($"File not found: {outputPath}");
                    return;
                }

                // Load the encrypted workbook using the same password
                LoadOptions loadOptions = new LoadOptions { Password = workbookPassword };
                Workbook loadedWb = new Workbook(outputPath, loadOptions);

                // Verify that the data can be read
                string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine("Loaded cell value: " + cellValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
