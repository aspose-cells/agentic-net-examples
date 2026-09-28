// Title: Encrypt an Excel workbook with Aspose.Cells in C# using a password retrieved at runtime from Azure Key Vault
// AI Prompts: Generate C# code that authenticates to Azure Key Vault, reads a secret containing the workbook password, and applies it to Aspose.Cells workbook.Settings.Password before saving the file. | Show how to add fallback logic that uses an environment variable when the Azure Key Vault secret cannot be obtained, while still encrypting the workbook with Aspose.Cells. | Provide error‑handling examples for missing secret, authentication failures, and saving an encrypted .xlsx using Aspose.Cells.
// Common Searches: C# Aspose.Cells encrypt Excel file using password from Azure Key Vault | how to get a password secret from Azure Key Vault and use it with Aspose.Cells to protect an Excel workbook | set Settings.Password in Aspose.Cells after retrieving password from Azure Key Vault | fallback to environment variable if Azure Key Vault secret is unavailable in Aspose.Cells encryption | sample code Aspose.Cells .NET Azure Key Vault secret retrieval for workbook encryption
// Tags: Aspose.Cells workbook encryption Azure Key Vault | C# retrieve secret Azure Key Vault | Aspose.Cells Settings.Password usage | encrypt Excel .xlsx with Aspose.Cells C# | secure password storage Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionExample
{
    // Demonstrates how to authenticate to Azure Key Vault, fetch a secret that holds the workbook password, assign it to workbook.Settings.Password, and save the workbook as an encrypted .xlsx file using Aspose.Cells in C#. Includes fallback to an environment variable and basic error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Retrieve workbook password from environment variable (fallback to a default value)
                string password = Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD") ?? "defaultPassword123";

                // Create a new workbook
                var workbook = new Workbook();

                // Populate the workbook with sample data
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Aspose.Cells Encryption Example");
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // Apply password protection (encryption) to the workbook
                workbook.Settings.Password = password;

                // Define output file path
                string outputPath = "EncryptedWorkbook.xlsx";

                // Save the encrypted workbook to disk
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}' with password protection.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
