// Title: Decrypt a password‑protected XLSX file on a background thread with Aspose.Cells for .NET and verify thread safety
// Common Searches: how to decrypt a password protected Excel workbook asynchronously using Aspose.Cells in C# | Aspose.Cells load encrypted XLSX with password on a Task and save unprotected | verify thread safety of Aspose.Cells workbook operations when using background threads | C# example for decrypting an encrypted Excel file on a separate thread with Aspose.Cells | using LoadOptions.Password to open encrypted workbook in a background task
// Tags: asynchronous workbook decryption Aspose.Cells | load encrypted xlsx with LoadOptions password | save decrypted workbook without password Aspose.Cells | thread‑safe Aspose.Cells workbook processing | background task Excel decryption C#

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample runs DecryptWorkbook on a Task, checks for the encrypted file, loads it with LoadOptions.Password, saves it unprotected, then the main thread waits, confirms the decrypted file exists, loads it without a password, reads cell A1 and prints the value, demonstrating background decryption and thread‑safe usage of Aspose.Cells.
class WorkbookDecryptionDemo
{
    // Path to the encrypted workbook
    private const string EncryptedFilePath = "encrypted.xlsx";
    // Password used to encrypt the workbook
    private const string Password = "myPassword";
    // Path where the decrypted workbook will be saved
    private const string DecryptedFilePath = "decrypted.xlsx";

    static void Main()
    {
        // Ensure Aspose.Cells license is set if you have one
        // License license = new License();
        // license.SetLicense("Aspose.Cells.NET.lic");

        // Run the decryption on a background thread
        Task decryptionTask = Task.Run(() => DecryptWorkbook());

        // Wait for the background operation to complete
        decryptionTask.Wait();

        // Verify that the decrypted file exists and can be opened without a password
        if (File.Exists(DecryptedFilePath))
        {
            try
            {
                // Load the decrypted workbook without any password
                Workbook decryptedWorkbook = new Workbook(DecryptedFilePath);
                // Access a cell to confirm the workbook is usable
                string firstCellValue = decryptedWorkbook.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine($"Decryption succeeded. First cell value: {firstCellValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open decrypted workbook: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("Decryption failed: decrypted file not found.");
        }
    }

    private static void DecryptWorkbook()
    {
        // Verify the encrypted file exists before attempting to load it
        if (!File.Exists(EncryptedFilePath))
        {
            Console.WriteLine($"Encrypted file not found: {EncryptedFilePath}");
            return;
        }

        try
        {
            // LoadOptions allows specifying the password for an encrypted workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = Password
            };

            // Load the encrypted workbook inside the background thread
            Workbook encryptedWorkbook = new Workbook(EncryptedFilePath, loadOptions);

            // Save the workbook without a password (i.e., decrypted)
            // By not providing a password in SaveOptions, the file will be saved unprotected
            encryptedWorkbook.Save(DecryptedFilePath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook decrypted and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during decryption: {ex.Message}");
        }
    }
}
