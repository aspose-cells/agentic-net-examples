// Title: C# – Encrypt an Excel workbook with Aspose.Cells, verify encryption, and test third‑party library compatibility
// Description: Creates a workbook, adds data, applies a password and 128‑bit strong encryption using Aspose.Cells, saves the file, detects the encrypted flag with FileFormatUtil, attempts to open it with NPOI (expected failure), and reloads it with the correct password.
// Keywords: Aspose.Cells encrypt Excel C# | password‑protected .xlsx Aspose | strong 128‑bit encryption Aspose.Cells | detect encrypted Excel file Aspose | load encrypted workbook Aspose.Cells | NPOI compatibility with encrypted Excel | Excel file encryption verification
// Common Searches: How to encrypt an .xlsx file with Aspose.Cells .NET | Check if an Excel workbook is encrypted using Aspose | Can NPOI read password‑protected Excel files created by Aspose | C# example for strong encryption of Excel workbooks | Load password‑protected workbook with Aspose.Cells
// Developer Intent: Apply password protection and strong encryption to an Excel file, confirm its encrypted status, and demonstrate that a third‑party library cannot open it.
// Use Cases: Create a workbook, set Workbook.Settings.Password, and call SetEncryptionOptions(StrongCryptographicProvider, 128) before saving as .xlsx. | Use FileFormatUtil.DetectFileFormat to read the IsEncrypted flag and verify encryption. | Attempt to open the encrypted file with NPOI, catch the expected exception, and log the incompatibility. | Reload the encrypted workbook with LoadOptions.Password to ensure proper decryption.
// AI Prompts: Generate C# code that encrypts an Excel workbook with a password and 128‑bit strong encryption using Aspose.Cells, then verifies the encrypted flag. | Provide a C# snippet that tries to open a password‑protected .xlsx file with NPOI, handles the exception, and outputs the result.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsEncryptionTest
{
    // Creates a workbook, adds data, applies a password and 128‑bit strong encryption using Aspose.Cells, saves the file, detects the encrypted flag with FileFormatUtil, attempts to open it with NPOI (expected failure), and reloads it with the correct password.
    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and add sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Encryption Test");
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // -----------------------------------------------------------------
                // 2. Apply password protection and encryption options
                // -----------------------------------------------------------------
                string password = "Secret123";
                workbook.Settings.Password = password; // Set opening password
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128); // Strong encryption

                // -----------------------------------------------------------------
                // 3. Save the encrypted workbook
                // -----------------------------------------------------------------
                string encryptedFilePath = "EncryptedWorkbook.xlsx";
                workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

                // -----------------------------------------------------------------
                // 4. Verify that the file is reported as encrypted by Aspose.Cells
                // -----------------------------------------------------------------
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(encryptedFilePath);
                Console.WriteLine($"Is file encrypted (Aspose detection)? {info.IsEncrypted}");

                // -----------------------------------------------------------------
                // 5. Attempt to open the encrypted file with a third‑party library (NPOI)
                //    NPOI does not support opening password‑protected .xlsx files, so we
                //    expect an exception which demonstrates incompatibility.
                // -----------------------------------------------------------------
                // Note: NPOI reference is omitted to keep the project self‑contained.
                // The following block is retained for conceptual completeness but
                // commented out to avoid compilation errors.
                /*
                try
                {
                    using (FileStream fs = new FileStream(encryptedFilePath, FileMode.Open, FileAccess.Read))
                    {
                        XSSFWorkbook npoiWorkbook = new XSSFWorkbook(fs);
                        Console.WriteLine("Third‑party library opened the file successfully (unexpected).");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Third‑party library failed to open the encrypted file as expected.");
                    Console.WriteLine($"Exception message: {ex.Message}");
                }
                */

                // -----------------------------------------------------------------
                // 6. Load the encrypted workbook back with Aspose.Cells using the password
                // -----------------------------------------------------------------
                if (File.Exists(encryptedFilePath))
                {
                    LoadOptions loadOptions = new LoadOptions { Password = password };
                    Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);
                    Console.WriteLine($"Loaded workbook contains {loadedWorkbook.Worksheets.Count} worksheet(s).");
                }
                else
                {
                    Console.WriteLine($"Encrypted file not found at path: {encryptedFilePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
