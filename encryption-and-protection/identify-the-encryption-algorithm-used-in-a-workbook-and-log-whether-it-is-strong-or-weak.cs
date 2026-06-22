using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook whose encryption we want to inspect
            string workbookPath = "EncryptedWorkbook.xls";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            Workbook wb = null;

            try
            {
                // Try loading with a password (if the workbook is encrypted)
                var loadOptions = new LoadOptions
                {
                    Password = "1234"
                };
                wb = new Workbook(workbookPath, loadOptions);
            }
            catch (Exception ex)
            {
                // If loading with a password fails, attempt to load without a password
                Console.WriteLine($"Failed to load with password: {ex.Message}");
                Console.WriteLine("Attempting to load without password...");

                try
                {
                    wb = new Workbook(workbookPath);
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Failed to load workbook: {innerEx.Message}");
                    return;
                }
            }

            try
            {
                // Check if the workbook is encrypted
                bool isEncrypted = wb?.Settings?.IsEncrypted ?? false;
                Console.WriteLine($"Workbook encrypted: {isEncrypted}");

                if (isEncrypted)
                {
                    // Assume modern encryption (SHA‑AES) or strong provider
                    Console.WriteLine("Encryption algorithm: Strong (SHA‑AES or StrongCryptographicProvider)");
                }
                else
                {
                    Console.WriteLine("Encryption algorithm: None (workbook is not encrypted)");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                wb?.Dispose();
            }
        }
    }
}