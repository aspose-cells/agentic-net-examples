// Title: Decrypt a password‑protected XLSX workbook on a background thread with Aspose.Cells for .NET and verify thread safety
// AI Prompts: Load an encrypted Excel file using LoadOptions.Password on a Task thread, modify a cell, and save it without a password. | Implement robust error handling for the decryption Task and surface exceptions to the calling thread. | Extend the sample to decrypt several password‑protected XLSX files in parallel using Task.WhenAll and validate each output.
// Common Searches: asp.net core decrypt password protected xlsx using Aspose.Cells in a background task | is Aspose.Cells thread safe when loading encrypted workbooks | c# async decryption of encrypted Excel file with Aspose.Cells LoadOptions | verify workbook content after decrypting Excel file on separate thread | process multiple encrypted Excel files concurrently with Aspose.Cells
// Tags: background thread workbook decryption Aspose.Cells | LoadOptions password property XLSX | thread‑safe Excel decryption .NET | verify decrypted workbook cell value | parallel XLSX decryption Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example checks for an encrypted XLSX file, runs a Task that loads the workbook with a password via LoadOptions, writes "Decrypted" to cell A1, saves it without a password, waits for the task, then loads the resulting file on the main thread to confirm the cell value, demonstrating thread‑safe decryption.
class Program
{
    static void Main()
    {
        // Paths and password for the encrypted workbook
        string encryptedPath = "encrypted.xlsx";
        string password = "myPassword";
        string decryptedPath = "decrypted.xlsx";

        // Ensure the encrypted file exists before attempting to load it
        if (!File.Exists(encryptedPath))
        {
            Console.WriteLine($"Error: Encrypted file \"{encryptedPath}\" not found.");
            return;
        }

        // Decrypt the workbook on a background thread
        Task decryptTask = Task.Run(() =>
        {
            try
            {
                // Load the encrypted workbook using the password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                Workbook wb = new Workbook(encryptedPath, loadOptions);

                // Simple operation to prove the workbook is usable after decryption
                wb.Worksheets[0].Cells["A1"].PutValue("Decrypted");

                // Save the workbook without a password (i.e., decrypted)
                wb.Save(decryptedPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Decryption failed: {ex.Message}");
                throw;
            }
        });

        try
        {
            // Wait for the background operation to finish
            decryptTask.Wait();
        }
        catch (AggregateException ae)
        {
            // Unwrap and display the original exception
            foreach (var inner in ae.InnerExceptions)
            {
                Console.WriteLine($"Error during decryption task: {inner.Message}");
            }
            return;
        }

        // Verify that the decrypted file was created
        if (!File.Exists(decryptedPath))
        {
            Console.WriteLine($"Error: Decrypted file \"{decryptedPath}\" was not created.");
            return;
        }

        // Verify on the main thread that the workbook was decrypted correctly
        try
        {
            Workbook verifyWb = new Workbook(decryptedPath);
            Console.WriteLine("Cell A1 value after decryption: " + verifyWb.Worksheets[0].Cells["A1"].StringValue);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Verification failed: {ex.Message}");
        }
    }
}
