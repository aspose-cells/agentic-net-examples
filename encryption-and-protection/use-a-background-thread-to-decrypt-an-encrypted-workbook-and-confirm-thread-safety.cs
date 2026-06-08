using System;
using System.IO;
using System.Text;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsThreadSafetyDemo
{
    class Program
    {
        // Path to the encrypted workbook
        private const string EncryptedFilePath = "encrypted.xlsx";
        // Password used for encryption/decryption
        private const string WorkbookPassword = "SecretPwd";

        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // Step 1: Create a workbook, add data, encrypt and save
                // -------------------------------------------------
                Workbook wbToEncrypt = new Workbook(); // create workbook
                Worksheet sheet = wbToEncrypt.Worksheets[0];

                // Populate some sample data
                for (int i = 0; i < 100; i++)
                {
                    sheet.Cells[i, 0].PutValue($"Row {i}");
                }

                // Set password to encrypt the workbook
                wbToEncrypt.Settings.Password = WorkbookPassword;

                // Save the encrypted workbook
                wbToEncrypt.Save(EncryptedFilePath);
                Console.WriteLine($"Encrypted workbook saved to '{EncryptedFilePath}'.");

                // -------------------------------------------------
                // Step 2: Decrypt the workbook on a background thread
                // -------------------------------------------------
                Thread decryptThread = new Thread(DecryptAndValidate);
                decryptThread.Start();
                decryptThread.Join(); // wait for background operation to finish

                Console.WriteLine("Background decryption and validation completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error in Main: {ex.Message}");
            }
        }

        // This method runs on a background thread
        private static void DecryptAndValidate()
        {
            try
            {
                // Verify that the encrypted file exists before attempting to load it
                if (!File.Exists(EncryptedFilePath))
                {
                    Console.WriteLine($"File not found: {EncryptedFilePath}");
                    return;
                }

                // LoadOptions with password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = WorkbookPassword
                };

                // Load the encrypted workbook
                Workbook decryptedWb = new Workbook(EncryptedFilePath, loadOptions);
                Console.WriteLine($"Workbook loaded on thread {Thread.CurrentThread.ManagedThreadId}.");

                // Verify that the workbook is no longer encrypted after successful load
                bool isEncrypted = decryptedWb.Settings.IsEncrypted;
                Console.WriteLine($"IsEncrypted after load: {isEncrypted}");

                // -------------------------------------------------
                // Step 3: Confirm thread‑safe reading of cell values
                // -------------------------------------------------
                Cells cells = decryptedWb.Worksheets[0].Cells;
                // Enable multi‑thread reading
                cells.MultiThreadReading = true;

                int totalRows = cells.MaxDataRow + 1; // number of rows with data
                int threadCount = 4;
                int rowsPerThread = totalRows / threadCount;
                int completedThreads = 0;
                StringBuilder errors = new StringBuilder();

                for (int t = 0; t < threadCount; t++)
                {
                    int startRow = t * rowsPerThread;
                    int endRow = (t == threadCount - 1) ? totalRows : startRow + rowsPerThread;

                    Thread reader = new Thread(() =>
                    {
                        try
                        {
                            for (int r = startRow; r < endRow; r++)
                            {
                                // Read cell value (thread‑safe because MultiThreadReading = true)
                                object value = cells[r, 0].Value;
                                Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId}: Cell[{r},0] = {value}");
                            }
                            Interlocked.Increment(ref completedThreads);
                        }
                        catch (Exception ex)
                        {
                            lock (errors)
                            {
                                errors.AppendLine($"Thread {Thread.CurrentThread.ManagedThreadId} error: {ex.Message}");
                            }
                        }
                    });
                    reader.Start();
                }

                // Wait until all reader threads finish
                while (completedThreads < threadCount)
                {
                    Thread.Sleep(100);
                }

                // Report any errors encountered during concurrent reads
                if (errors.Length > 0)
                {
                    Console.WriteLine("Errors occurred during multi‑thread reading:");
                    Console.WriteLine(errors.ToString());
                }
                else
                {
                    Console.WriteLine("All reader threads completed successfully without errors.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in DecryptAndValidate: {ex.Message}");
            }
        }
    }
}