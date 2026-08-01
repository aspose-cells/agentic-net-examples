// Title: Decrypt a password‑protected Excel workbook on a background thread with Aspose.Cells for .NET
// Description: This example creates a workbook, encrypts it with Settings.Password, saves it, then loads the file on a separate Thread using LoadOptions.Password. It checks Settings.IsEncrypted, reads a cell to confirm data integrity, captures any exception, disposes resources, and deletes the temporary file, demonstrating thread‑safe decryption in C#.
// Keywords: Aspose.Cells | C# | .NET | encrypted workbook | password protected Excel | background thread | thread safety | LoadOptions.Password | decryption example | UI responsive loading | global | US | Europe
// Common Searches: Aspose.Cells load encrypted Excel on a worker thread | Is Aspose.Cells thread safe for password decryption | C# background thread decrypt Excel file Aspose | How to use LoadOptions.Password with Aspose.Cells | Prevent UI blocking when opening protected workbook
// Developer Intent: Load and decrypt a password‑protected Excel file on a background thread while verifying that Aspose.Cells operations remain thread‑safe.
// Use Cases: Open large protected workbooks without freezing the UI. | Validate that Settings.IsEncrypted reflects the correct state after decryption in a worker thread. | Ensure proper disposal of workbook objects loaded on background threads to avoid memory leaks.
// AI Prompts: Generate a C# Task‑based version of this example that returns the decrypted Workbook object. | Explain how to propagate exceptions from a background thread when loading an encrypted workbook with Aspose.Cells. | Show best practices for disposing an Aspose.Cells Workbook loaded on a separate thread to prevent cross‑thread resource conflicts.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    // This example creates a workbook, encrypts it with Settings.Password, saves it, then loads the file on a separate Thread using LoadOptions.Password. It checks Settings.IsEncrypted, reads a cell to confirm data integrity, captures any exception, disposes resources, and deletes the temporary file, demonstrating thread‑safe decryption in C#.
    class Program
    {
        // Path for the temporary encrypted workbook
        private const string EncryptedFilePath = "encrypted_demo.xlsx";
        // Password used for encryption and decryption
        private const string WorkbookPassword = "SecretPwd";

        static void Main()
        {
            // -------------------------------------------------
            // Step 1: Create a workbook, add data, encrypt and save
            // -------------------------------------------------
            // Create a new workbook (rule: create)
            Workbook wb = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Thread safety test");

            // Set password to encrypt the workbook (rule: property)
            wb.Settings.Password = WorkbookPassword;

            // Save the encrypted workbook (rule: save)
            wb.Save(EncryptedFilePath);

            // Dispose the original workbook
            wb.Dispose();

            // -------------------------------------------------
            // Step 2: Load the encrypted workbook on a background thread
            // -------------------------------------------------
            // Variable to hold the loaded workbook reference
            Workbook loadedWorkbook = null;
            // Variable to capture any exception from the thread
            Exception threadException = null;

            Thread loadThread = new Thread(() =>
            {
                try
                {
                    // Prepare load options with the password (rule: load)
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = WorkbookPassword
                    };

                    // Load the encrypted workbook using the password
                    loadedWorkbook = new Workbook(EncryptedFilePath, loadOptions);

                    // Verify that the workbook reports as encrypted
                    bool isEncrypted = loadedWorkbook.Settings.IsEncrypted;
                    Console.WriteLine($"[Thread] Workbook Settings.IsEncrypted: {isEncrypted}");

                    // Read the cell value to ensure data integrity
                    string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
                    Console.WriteLine($"[Thread] Cell A1 value: {cellValue}");
                }
                catch (Exception ex)
                {
                    // Capture exception for reporting after the thread joins
                    threadException = ex;
                }
            });

            // Start the background thread
            loadThread.Start();

            // Wait for the thread to finish
            loadThread.Join();

            // -------------------------------------------------
            // Step 3: Confirm thread safety outcome
            // -------------------------------------------------
            if (threadException != null)
            {
                Console.WriteLine($"Error during background load: {threadException.Message}");
            }
            else
            {
                Console.WriteLine("Background decryption completed successfully.");
            }

            // Clean up the loaded workbook
            loadedWorkbook?.Dispose();

            // Optionally delete the temporary file
            if (File.Exists(EncryptedFilePath))
            {
                File.Delete(EncryptedFilePath);
            }
        }
    }
}
