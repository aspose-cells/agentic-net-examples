// Title: Decrypt a Password‑Protected Aspose.Cells Workbook on a Background Thread (C#)
// Description: This example creates a workbook, applies a password and strong AES encryption, saves it, then loads the file on a background thread using LoadOptions. It verifies that the workbook is no longer encrypted, enables MultiThreadReading for safe concurrent access, reads a cell value to confirm successful decryption, and cleans up—all while demonstrating thread‑safe usage of Aspose.Cells in .NET.
// Keywords: Aspose.Cells | C# | decrypt workbook | background thread | password protected Excel | strong encryption | LoadOptions | MultiThreadReading | .NET | thread safety
// Common Searches: Aspose.Cells load encrypted Excel on separate thread | C# decrypt password protected workbook using Aspose.Cells | Is Aspose.Cells workbook loading thread safe | Enable MultiThreadReading after opening encrypted file | How to set strong encryption with Aspose.Cells
// Developer Intent: Open an encrypted Excel file in a background thread, confirm decryption, and ensure safe concurrent reads.
// Use Cases: Keep the UI responsive by decrypting large protected workbooks off the UI thread. | Validate that a workbook is fully decrypted before any processing begins. | Enable multi‑threaded cell reads after opening a password‑protected file. | Integrate secure workbook handling into server‑side batch jobs.
// AI Prompts: Generate C# code that uses Task.Run to open an encrypted workbook with Aspose.Cells and returns the first cell value. | Explain Aspose.Cells thread‑safety model for reading cells after a workbook is loaded with a password. | List exceptions thrown by Aspose.Cells when an incorrect password is supplied during LoadOptions. | Show how to configure Aspose.Cells to use AES‑256 encryption when saving a workbook.

using System;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    // This example creates a workbook, applies a password and strong AES encryption, saves it, then loads the file on a background thread using LoadOptions. It verifies that the workbook is no longer encrypted, enables MultiThreadReading for safe concurrent access, reads a cell value to confirm successful decryption, and cleans up—all while demonstrating thread‑safe usage of Aspose.Cells in .NET.
    class Program
    {
        // Path for the temporary encrypted workbook
        private const string EncryptedFilePath = "encrypted_demo.xlsx";

        static void Main()
        {
            // -------------------------------------------------
            // Step 1: Create a workbook, add data and encrypt it
            // -------------------------------------------------
            Workbook wb = new Workbook();                         // create workbook
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Thread safety test");    // add sample data

            // Set password to protect the workbook
            wb.Settings.Password = "SecretPwd";

            // Optionally set strong encryption (requires Aspose.Cells 23.5+)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            wb.Save(EncryptedFilePath);
            Console.WriteLine($"Encrypted workbook saved to '{EncryptedFilePath}'.");

            // -------------------------------------------------
            // Step 2: Decrypt the workbook on a background thread
            // -------------------------------------------------
            Thread decryptThread = new Thread(() =>
            {
                try
                {
                    // LoadOptions with password to open the encrypted file
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = "SecretPwd"
                    };

                    // Load the workbook (decryption happens internally)
                    Workbook loadedWb = new Workbook(EncryptedFilePath, loadOptions);

                    // Verify that the workbook is no longer encrypted after loading
                    bool isEncrypted = loadedWb.Settings.IsEncrypted;
                    Console.WriteLine($"[Thread] Workbook IsEncrypted after load: {isEncrypted}");

                    // Enable multi‑thread reading for safety (not strictly required here)
                    loadedWb.Worksheets[0].Cells.MultiThreadReading = true;

                    // Read the cell value to confirm successful decryption
                    string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
                    Console.WriteLine($"[Thread] Decrypted cell value: {cellValue}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Thread] Exception: {ex.Message}");
                }
            });

            // Start the background thread
            decryptThread.IsBackground = true;
            decryptThread.Start();

            // Wait for the thread to finish
            decryptThread.Join();

            // Cleanup
            wb.Dispose();
            Console.WriteLine("Demo completed.");
        }
    }
}
