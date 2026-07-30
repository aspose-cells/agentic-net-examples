// Title: Decrypt a Password‑Protected Aspose.Cells Workbook on a Background Thread (C#)
// Description: This example shows how to create an Excel workbook, encrypt it with a password, save it, and then load and decrypt it on a separate background thread using Aspose.Cells LoadOptions. It verifies that Settings.IsEncrypted is false after loading and reads a cell value to confirm successful decryption, demonstrating thread‑safe workbook loading in .NET.
// Keywords: Aspose.Cells | C# | decrypt encrypted workbook | password protected Excel | background thread | thread safety | LoadOptions password | multithreaded Excel processing | Excel encryption Aspose | concurrent workbook loading
// Common Searches: How to open a password‑protected Excel file with Aspose.Cells in a background thread | Is Aspose.Cells workbook loading thread‑safe? | C# example for decrypting an encrypted workbook on a separate thread | LoadOptions password Aspose.Cells multithreaded | Decrypt Excel file without blocking UI using Aspose.Cells
// Developer Intent: Load and decrypt a password‑protected Excel workbook on a background thread to verify that Aspose.Cells operations are safe for multithreaded scenarios.
// Use Cases: Process encrypted reports in parallel background workers to keep the UI responsive. | Decrypt user‑uploaded protected spreadsheets on a server‑side thread before data extraction. | Validate workbook integrity after decryption in a multithreaded import pipeline.
// AI Prompts: Generate C# code that opens an encrypted Excel file with Aspose.Cells using Task.Run, includes exception handling, and logs thread‑safety verification. | Write an MSTest unit test that loads an encrypted workbook on a background thread and asserts Settings.IsEncrypted is false and cell values match the original. | Explain how to configure Aspose.Cells for concurrent loading of multiple password‑protected workbooks in a .NET Core application.

using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    // This example shows how to create an Excel workbook, encrypt it with a password, save it, and then load and decrypt it on a separate background thread using Aspose.Cells LoadOptions. It verifies that Settings.IsEncrypted is false after loading and reads a cell value to confirm successful decryption, demonstrating thread‑safe workbook loading in .NET.
    class Program
    {
        // Path for the temporary encrypted workbook
        private const string EncryptedFilePath = "encrypted_demo.xlsx";
        // Password used for encryption/decryption
        private const string WorkbookPassword = "SecretPwd";

        static void Main()
        {
            // -------------------------------------------------
            // Step 1: Create a workbook, add data, encrypt & save
            // -------------------------------------------------
            // Create a new workbook (uses the provided creation rule)
            Workbook wb = new Workbook();

            // Add a sample value to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Thread safety test");

            // Encrypt the workbook by setting a password (uses the Settings.Password property)
            wb.Settings.Password = WorkbookPassword;

            // Save the encrypted workbook (uses the provided save rule)
            wb.Save(EncryptedFilePath);

            // Dispose the original workbook
            wb.Dispose();

            // -------------------------------------------------
            // Step 2: Decrypt the workbook on a background thread
            // -------------------------------------------------
            Thread decryptThread = new Thread(() =>
            {
                // LoadOptions with the password to open the encrypted file (uses the load rule)
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = WorkbookPassword
                };

                // Load the encrypted workbook (constructor with file path and LoadOptions)
                Workbook loadedWb = new Workbook(EncryptedFilePath, loadOptions);

                // Verify that the workbook is no longer considered encrypted after loading
                bool isEncrypted = loadedWb.Settings.IsEncrypted;
                Console.WriteLine($"[Background Thread] Workbook.IsEncrypted after load: {isEncrypted}");

                // Read the previously stored cell value to confirm successful decryption
                string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine($"[Background Thread] Decrypted cell value: {cellValue}");

                // Clean up
                loadedWb.Dispose();
            });

            // Start the background thread
            decryptThread.Start();

            // Wait for the background thread to finish
            decryptThread.Join();

            // -------------------------------------------------
            // Step 3: Clean up the temporary file (optional)
            // -------------------------------------------------
            if (System.IO.File.Exists(EncryptedFilePath))
            {
                System.IO.File.Delete(EncryptedFilePath);
            }

            Console.WriteLine("Demo completed.");
        }
    }
}
