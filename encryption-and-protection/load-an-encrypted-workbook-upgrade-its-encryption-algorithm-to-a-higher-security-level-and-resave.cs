// Title: C# – Upgrade an Encrypted Excel Workbook to 256‑bit StrongCryptographicProvider with Aspose.Cells
// Description: Shows how to open a password‑protected .xlsx using Aspose.Cells LoadOptions, switch its encryption to the 256‑bit StrongCryptographicProvider via SetEncryptionOptions, and save the workbook under a new name. The sample also creates a test file if none exists and includes comprehensive error handling.
// Keywords: Aspose.Cells | C# encryption upgrade | Excel 256‑bit encryption | StrongCryptographicProvider | SetEncryptionOptions | LoadOptions password | re‑encrypt workbook .NET | upgrade Excel security | encrypted workbook sample | Aspose.Cells example
// Common Searches: How to change encryption type of an existing Excel file with Aspose.Cells | Upgrade Excel workbook encryption to 256‑bit using C# | Re‑encrypt password‑protected workbook Aspose.Cells .NET | SetEncryptionOptions StrongCryptographicProvider example | Load encrypted Excel file with password Aspose.Cells
// Developer Intent: Replace the encryption algorithm of a password‑protected Excel workbook with a stronger 256‑bit scheme and write the upgraded file.
// Use Cases: Bring legacy encrypted reports up to current security standards (e.g., ISO 27001, GDPR) by re‑saving them with 256‑bit encryption. | Secure user‑uploaded spreadsheets before storing them in a cloud repository or document management system. | Automate batch processing that loads encrypted workbooks, upgrades their encryption, and outputs the refreshed files to a designated folder.
// AI Prompts: Generate C# code that opens an encrypted Excel file with a given password, upgrades its encryption to StrongCryptographicProvider 256‑bit using Aspose.Cells, and saves it as a new file. | Explain the parameters of SetEncryptionOptions in Aspose.Cells and list all supported EncryptionType values. | Provide best‑practice error handling for workbook encryption upgrades in a .NET application, including file‑not‑found and Aspose.Cells exceptions.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to open a password‑protected .xlsx using Aspose.Cells LoadOptions, switch its encryption to the 256‑bit StrongCryptographicProvider via SetEncryptionOptions, and save the workbook under a new name. The sample also creates a test file if none exists and includes comprehensive error handling.
class UpgradeEncryptionDemo
{
    static void Main()
    {
        // Paths for the input (existing encrypted) and output workbooks
        string inputPath = "EncryptedWorkbook.xlsx";
        string outputPath = "UpgradedEncryptedWorkbook.xlsx";

        // Password used to open the existing encrypted workbook
        string currentPassword = "oldPassword";

        try
        {
            // If the input file does not exist, create a sample encrypted workbook first
            if (!File.Exists(inputPath))
            {
                // Create a simple workbook
                Workbook sample = new Workbook();
                sample.Worksheets[0].Cells["A1"].PutValue("Sample data");

                // Apply password protection (default encryption will be used)
                sample.Settings.Password = currentPassword;
                sample.Save(inputPath);
                Console.WriteLine($"Sample encrypted workbook created at '{inputPath}'.");
            }

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = currentPassword
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Keep the same password for the upgraded file
            workbook.Settings.Password = currentPassword;

            // Upgrade encryption algorithm to a stronger one (256‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook with upgraded encryption
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved with upgraded encryption at '{outputPath}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.Message}");
        }
        catch (CellsException cellsEx)
        {
            Console.WriteLine($"Aspose.Cells error: {cellsEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
