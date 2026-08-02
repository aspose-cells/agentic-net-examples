// Title: Decrypt a password‑protected Excel file with Aspose.Cells for .NET and verify data integrity
// Description: This C# example shows how to load an encrypted .xlsx using LoadOptions.Password, read a cell to capture the original value, clear the workbook password, save the file unencrypted, reload it, and compare the values to confirm that the data remained unchanged after decryption.
// Keywords: Aspose.Cells decrypt workbook | C# remove Excel password | load encrypted .xlsx Aspose | verify Excel data integrity | unprotect workbook Aspose.Cells | LoadOptions.Password | IsEncrypted flag | Aspose.Cells .NET example
// Common Searches: Aspose.Cells open password protected Excel file | C# remove password from .xlsx using Aspose | How to check if Excel workbook is encrypted with Aspose.Cells | Validate data after decrypting Excel with Aspose | Save unencrypted workbook after decryption Aspose.Cells
// Developer Intent: Open an encrypted Excel workbook with the correct password, strip its protection, save it as an unencrypted file, and ensure the original cell content is unchanged.
// Use Cases: Load an encrypted workbook by supplying LoadOptions.Password and read a specific cell. | Clear Settings.Password to remove encryption, then save the workbook to a new location. | Reload the saved file without a password and compare cell values to confirm data integrity.
// AI Prompts: Write C# code that uses Aspose.Cells to open a password‑protected workbook, remove the password, save it unencrypted, and verify that cell A1 still contains the original text. | Provide error‑handling best practices for decrypting an Excel file with Aspose.Cells, covering FileNotFoundException, CellsException, and generic exceptions. | Explain how to use the IsEncrypted property before and after clearing Settings.Password to confirm that encryption has been removed.

using System;
using System.IO;
using Aspose.Cells;

// This C# example shows how to load an encrypted .xlsx using LoadOptions.Password, read a cell to capture the original value, clear the workbook password, save the file unencrypted, reload it, and compare the values to confirm that the data remained unchanged after decryption.
class DecryptWorkbookDemo
{
    static void Main()
    {
        // Paths and password
        string encryptedPath = "encrypted.xlsx";
        string unprotectedPath = "unprotected.xlsx";
        string password = "mySecret";

        try
        {
            // Ensure the encrypted workbook exists; if not, create and encrypt it
            if (!File.Exists(encryptedPath))
            {
                // Create a simple workbook with a value in A1
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Secret Data");

                // Encrypt the workbook with the password
                tempWb.Settings.Password = password;
                tempWb.Save(encryptedPath);
                Console.WriteLine($"Created encrypted workbook: {encryptedPath}");
            }

            // Load the encrypted workbook using the correct password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedPath, loadOptions);

            // Verify encryption status
            Console.WriteLine("Workbook IsEncrypted: " + workbook.Settings.IsEncrypted);

            // Read a cell value to confirm data integrity before decryption
            string originalValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Original A1 value: " + originalValue);

            // Remove encryption by clearing the password
            workbook.Settings.Password = null;
            workbook.Save(unprotectedPath);
            Console.WriteLine($"Saved unprotected workbook: {unprotectedPath}");

            // Reload the saved workbook without providing a password
            Workbook unprotectedWorkbook = new Workbook(unprotectedPath);
            string reloadedValue = unprotectedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Reloaded A1 value: " + reloadedValue);

            // Confirm that the data remained unchanged after decryption
            bool dataIntact = originalValue == reloadedValue;
            Console.WriteLine("Data integrity after decryption: " + dataIntact);
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine("File not found: " + fnfEx.Message);
        }
        catch (CellsException cellsEx)
        {
            Console.WriteLine("Aspose.Cells error: " + cellsEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
