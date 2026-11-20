// Title: Decrypt a password‑protected Excel workbook on a background thread with Aspose.Cells for .NET and verify thread‑safe operations
// Common Searches: how to load an encrypted .xlsx file with a password using Aspose.Cells in C# | run Aspose.Cells workbook calculations on a Task thread safely | save a decrypted Excel workbook without a password after background processing in .NET | verify thread safety when modifying a workbook loaded with LoadOptions in Aspose.Cells | C# example for decrypting and editing a protected workbook on a separate thread
// Tags: load encrypted workbook with LoadOptions Aspose.Cells | background thread workbook processing Aspose.Cells | thread‑safe formula calculation Aspose.Cells | save decrypted Excel file Aspose.Cells | C# Aspose.Cells password protected workbook

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace WorkbookDecryptionDemoApp
{
    // The sample loads a password‑protected XLSX file using LoadOptions, runs workbook operations (formula recalculation, cell update) inside a Task.Run background thread, saves the workbook without a password, and then reloads the saved file on the main thread to confirm successful decryption and thread‑safe processing.
    class WorkbookDecryptionDemo
    {
        // Path to the encrypted workbook file
        private const string EncryptedFilePath = @"C:\Temp\EncryptedWorkbook.xlsx";

        // Password used to encrypt the workbook
        private const string WorkbookPassword = "MySecretPassword";

        // Path to save the decrypted workbook after processing
        private const string DecryptedFilePath = @"C:\Temp\DecryptedWorkbook.xlsx";

        static void Main()
        {
            try
            {
                // Verify that the encrypted file exists
                if (!File.Exists(EncryptedFilePath))
                {
                    Console.WriteLine($"Encrypted file not found: {EncryptedFilePath}");
                    return;
                }

                // Load the encrypted workbook (required to obtain the Workbook object)
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = WorkbookPassword
                };
                var workbook = new Workbook(EncryptedFilePath, loadOptions);

                // Perform workbook operations on a background thread
                Task backgroundTask = Task.Run(() =>
                {
                    try
                    {
                        Console.WriteLine($"Background thread: Workbook loaded with {workbook.Worksheets.Count} worksheets.");

                        // Recalculate all formulas (thread‑safe)
                        workbook.CalculateFormula();

                        // Modify a cell value
                        Worksheet sheet = workbook.Worksheets[0];
                        sheet.Cells["A1"].PutValue("Decrypted on background thread");

                        // Ensure the output directory exists
                        string outputDir = Path.GetDirectoryName(DecryptedFilePath);
                        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                        {
                            Directory.CreateDirectory(outputDir);
                        }

                        // Save the workbook without a password
                        workbook.Save(DecryptedFilePath, SaveFormat.Xlsx);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Background thread error: {ex.Message}");
                    }
                });

                // Wait for the background operation to finish
                backgroundTask.Wait();

                // Verify that the decrypted file was created
                if (File.Exists(DecryptedFilePath))
                {
                    // Load the saved workbook to confirm decryption succeeded
                    var decryptedWorkbook = new Workbook(DecryptedFilePath);
                    Console.WriteLine($"Main thread: Decrypted workbook loaded with {decryptedWorkbook.Worksheets.Count} worksheets.");
                    Console.WriteLine($"Cell A1 value: {decryptedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
                }
                else
                {
                    Console.WriteLine("Main thread: Decrypted file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
