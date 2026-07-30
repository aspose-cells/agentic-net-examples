// Title: Detect Encrypted OOXML (.xlsx) and Decrypt with Aspose.Cells for .NET
// Description: C# sample that uses Aspose.Cells to identify an encrypted Excel workbook, verify the supplied password, load the file with the password, and save an unprotected copy to guarantee compatibility. Includes robust error handling for CellsException and generic failures.
// Keywords: Aspose.Cells detect encrypted xlsx | verify Excel password Aspose | load encrypted workbook .NET | decrypt OOXML file | save unencrypted Excel copy | FileFormatUtil DetectFileFormat | Aspose.Cells LoadOptions password
// Common Searches: How to check if an .xlsx is password protected with Aspose.Cells | Verify password before opening encrypted Excel in C# | Decrypt an encrypted OOXML spreadsheet using Aspose.Cells | Load and save unprotected copy of a protected workbook | Aspose.Cells example for encrypted file detection
// Developer Intent: Identify encryption status, confirm password validity, open the protected workbook, and export a decrypted version.
// Use Cases: Pre‑flight check for encryption to avoid runtime errors when processing Excel files. | Password validation without fully loading the workbook, useful for authentication workflows. | Conversion of a secured spreadsheet to an unprotected file for downstream analytics or reporting.
// AI Prompts: Write C# code that uses Aspose.Cells to detect if an Excel file is encrypted and validate a given password. | Show how to load an encrypted .xlsx with a password and save it as a plain workbook using Aspose.Cells for .NET. | Explain best practices for handling CellsException when working with password‑protected workbooks in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# sample that uses Aspose.Cells to identify an encrypted Excel workbook, verify the supplied password, load the file with the password, and save an unprotected copy to guarantee compatibility. Includes robust error handling for CellsException and generic failures.
    class DetectAndDecryptEncryptedOOXML
    {
        static void Main()
        {
            try
            {
                // Path to the encrypted OOXML file (e.g., .xlsx)
                string filePath = "encrypted.xlsx";

                // Ensure the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File \"{filePath}\" not found.");
                    return;
                }

                // Password used to protect the file
                string password = "test";

                // -------------------------------------------------
                // 1. Detect file format and encryption status
                // -------------------------------------------------
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

                // -------------------------------------------------
                // 2. Verify that the supplied password is correct (if encrypted)
                // -------------------------------------------------
                bool passwordValid = true;
                if (formatInfo.IsEncrypted)
                {
                    using (FileStream stream = File.OpenRead(filePath))
                    {
                        passwordValid = FileFormatUtil.VerifyPassword(stream, password);
                    }
                    Console.WriteLine($"Password valid: {passwordValid}");
                }

                if (!passwordValid)
                {
                    Console.WriteLine("The provided password is incorrect. Cannot load the workbook.");
                    return;
                }

                // -------------------------------------------------
                // 3. Load the workbook with the correct password
                // -------------------------------------------------
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully after decryption.");

                // -------------------------------------------------
                // 4. Save the workbook to a new file to ensure compatibility
                // -------------------------------------------------
                string outputPath = "decrypted_copy.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Decrypted workbook saved to: {outputPath}");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
