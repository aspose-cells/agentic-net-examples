// Title: Change an Encrypted Excel Workbook Password Using Aspose.Cells and a Secure Key Vault (C#)
// Description: Shows how to open a password‑protected .xlsx file by retrieving the old password from a key‑vault (e.g., Azure Key Vault or AWS Secrets Manager), replace it with a new password, and save the workbook with Aspose.Cells. The sample includes file‑existence validation and handling of Aspose.Cells exceptions.
// Keywords: Aspose.Cells | C# Excel password change | encrypted workbook | load workbook with password | save workbook with new password | Azure Key Vault | AWS Secrets Manager | secure secret retrieval | Excel file protection | LoadOptions.Password | Workbook.Settings.Password
// Common Searches: How to change password of an encrypted Excel file using Aspose.Cells C# | Aspose.Cells load encrypted workbook with old password | Save Excel workbook with new password Aspose.Cells | Retrieve Excel passwords from Azure Key Vault in .NET | Replace workbook password without recreating file Aspose.Cells
// Developer Intent: Replace the existing password of an encrypted Excel workbook with a new one obtained from a secure key vault and save the updated file.
// Use Cases: Fetch old and new passwords from Azure Key Vault, AWS Secrets Manager, or another secret store, then open the protected .xlsx with LoadOptions.Password. | Assign the new password to workbook.Settings.Password and save to a different file name to avoid overwriting the original. | Validate that the source file exists before loading and catch CellsException for incorrect passwords or corrupted workbooks. | Swap the placeholder GetSecret method with real SDK calls to keep credentials out of source code.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an encrypted Excel workbook with a password retrieved from Azure Key Vault, change the password, and save the file. | Provide an example of integrating AWS Secrets Manager with Aspose.Cells to replace an Excel workbook’s password in a .NET console application. | Explain how to handle CellsException when the supplied old password is invalid while changing the workbook password using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Shows how to open a password‑protected .xlsx file by retrieving the old password from a key‑vault (e.g., Azure Key Vault or AWS Secrets Manager), replace it with a new password, and save the workbook with Aspose.Cells. The sample includes file‑existence validation and handling of Aspose.Cells exceptions.
    class ChangeWorkbookPassword
    {
        static void Main()
        {
            try
            {
                // Retrieve passwords from a secure key vault (placeholder implementation)
                string oldPassword = GetSecret("OldWorkbookPassword");
                string newPassword = GetSecret("NewWorkbookPassword");

                string inputPath = "encryptedWorkbook.xlsx";
                string outputPath = "encryptedWorkbook_updated.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the encrypted workbook using the old password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Change the workbook password
                workbook.Settings.Password = newPassword;

                // Save the workbook with the new password
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
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

        // Placeholder method simulating secure key vault retrieval
        static string GetSecret(string secretName)
        {
            // Replace this with actual key vault integration (e.g., Azure Key Vault, AWS Secrets Manager)
            return secretName switch
            {
                "OldWorkbookPassword" => "oldPass123",
                "NewWorkbookPassword" => "newPass456",
                _ => string.Empty,
            };
        }
    }
}
