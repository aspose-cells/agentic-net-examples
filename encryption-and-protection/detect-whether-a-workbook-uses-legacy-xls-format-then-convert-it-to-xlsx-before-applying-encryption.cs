// Title: Detect XLS format, convert to XLSX and apply password protection with Aspose.Cells for .NET
// Description: Demonstrates using Aspose.Cells to identify a legacy XLS workbook, convert it to XLSX, set a password through Workbook.Settings.Password, save the encrypted file, and optionally confirm encryption by re‑loading with LoadOptions.
// Keywords: Aspose.Cells detect file format | XLS to XLSX conversion C# | Excel password protection .NET | FileFormatUtil | Workbook.Settings.Password | Encrypt Excel workbook Aspose | Legacy Excel conversion | Secure XLSX output
// Common Searches: how to detect xls vs xlsx with Aspose.Cells | convert old xls to xlsx and set password c# | asp.net encrypt excel file after conversion | verify excel file encryption Aspose.Cells | batch convert and protect legacy excel files
// Developer Intent: Identify workbook type, convert legacy XLS to XLSX, and protect the result with a password.
// Use Cases: Automated pipeline that normalizes mixed‑format spreadsheets to XLSX and secures them before storage. | Web service that receives user‑uploaded Excel files, ensures they are saved as encrypted XLSX, and returns a download link. | Compliance workflow that validates encryption by re‑opening the saved file with the supplied password.
// AI Prompts: Generate C# code using Aspose.Cells to detect an Excel file's format, convert XLS to XLSX if needed, and encrypt it with a password. | Explain how to confirm that a workbook saved with Aspose.Cells is encrypted and how to handle decryption failures. | Provide best‑practice recommendations for batch processing, converting, and password‑protecting large numbers of legacy Excel files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates using Aspose.Cells to identify a legacy XLS workbook, convert it to XLSX, set a password through Workbook.Settings.Password, save the encrypted file, and optionally confirm encryption by re‑loading with LoadOptions.
    public class ConvertLegacyXlsAndEncrypt
    {
        /// <param name="sourcePath">Path to the original workbook (XLS or XLSX).</param>
        /// <param name="outputPath">Path where the encrypted XLSX workbook will be saved.</param>
        /// <param name="password">Password to protect the workbook.</param>
        public static void Run(string sourcePath, string outputPath, string password)
        {
            try
            {
                // Ensure the source file exists to avoid FileNotFoundException.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Detect the file format of the source workbook.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);

                // Load the workbook (Aspose.Cells automatically selects the correct load format).
                Workbook workbook = new Workbook(sourcePath);

                // Apply password protection (encryption) to the workbook.
                workbook.Settings.Password = password;

                // Save as XLSX. If the source is a legacy XLS, this also performs the conversion.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Optional: verify that the saved file is encrypted.
                LoadOptions loadOptions = new LoadOptions { Password = password };
                Workbook encryptedWorkbook = new Workbook(outputPath, loadOptions);
                Console.WriteLine($"Workbook encrypted: {encryptedWorkbook.Settings.IsEncrypted}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the console application.
        public static void Main(string[] args)
        {
            // Example usage – replace with actual paths and password as needed.
            string sourcePath = "input.xls";
            string outputPath = "encrypted_output.xlsx";
            string password = "StrongPassword123";

            Run(sourcePath, outputPath, password);
        }
    }
}
