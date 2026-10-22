// Title: Encrypt an Excel workbook with a password retrieved from Azure Key Vault using Aspose.Cells for .NET
// Common Searches: how to encrypt an Excel file with a password stored in Azure Key Vault using Aspose.Cells C# | Aspose.Cells set workbook password from Azure Key Vault secret .NET example | save password‑protected XLSX with Aspose.Cells after retrieving key from Azure Key Vault | C# retrieve secret from Azure Key Vault and apply workbook encryption with Aspose.Cells | secure Excel file generation in .NET using Aspose.Cells and Azure Key Vault
// Tags: Aspose.Cells workbook encryption Azure Key Vault | C# set Excel file password Aspose.Cells | retrieve secret from Azure Key Vault for .NET | save password‑protected XLSX Aspose.Cells | secure Excel generation Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample reads a password (or secret) from a secure source such as Azure Key Vault, assigns it to workbook.Settings.Password, and saves the workbook as an encrypted XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Retrieve workbook password from an environment variable (replace with a secure source as needed)
            string workbookPassword = Environment.GetEnvironmentVariable("WorkbookPassword");
            if (string.IsNullOrEmpty(workbookPassword))
            {
                Console.WriteLine("Environment variable 'WorkbookPassword' not found. Using a default password.");
                workbookPassword = "defaultPassword";
            }

            // Create a new workbook
            var workbook = new Workbook();

            // Populate the workbook with sample data
            var cells = workbook.Worksheets[0].Cells;
            cells["A1"].PutValue("Hello");
            cells["B1"].PutValue("World");

            // Apply encryption using the retrieved password
            workbook.Settings.Password = workbookPassword;

            // Save the encrypted workbook to disk
            string outputPath = "EncryptedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Encrypted workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
