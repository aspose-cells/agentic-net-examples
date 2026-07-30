// Title: C# Batch Decrypt Encrypted Excel Workbooks with Aspose.Cells – Sample Code
// Description: A concise C# example that loops through a dictionary of Excel file paths and passwords, checks each file's existence, detects its format, loads the workbook via Aspose.Cells LoadOptions, clears workbook‑level encryption, optionally unprotects protected sheets, and writes a decrypted copy with a "_decrypted" suffix. Includes robust error handling and logging for missing files and format detection.
// Keywords: Aspose.Cells | C# decrypt Excel files | remove Excel password programmatically | batch Excel decryption | load encrypted workbook Aspose | unprotect Excel sheets C# | sample code GitHub | US developers | European developers
// Common Searches: batch decrypt Excel files C# Aspose.Cells | remove password from multiple .xlsx using code | Aspose.Cells load encrypted workbook with password | how to unprotect Excel workbook programmatically | C# example for decrypting encrypted spreadsheets
// Developer Intent: Programmatically open several password‑protected Excel workbooks, strip encryption/protection, and save unencrypted versions.
// Use Cases: Automate nightly decryption of incoming encrypted reports before importing data into a database. | Create decrypted copies of confidential templates so internal teams can edit them. | Process a mixed folder of .xls and .xlsx files, each with its own password, for compliance audits.
// AI Prompts: Write C# code that reads a Dictionary<string,string> of Excel paths and passwords, opens each workbook with Aspose.Cells, removes encryption and protection, and saves the file with a '_decrypted' suffix. | Explain how to differentiate between file encryption password and worksheet protection password when using Aspose.Cells. | Suggest robust logging and exception‑handling strategies for a batch Excel decryption utility, including handling missing files and format detection.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// A concise C# example that loops through a dictionary of Excel file paths and passwords, checks each file's existence, detects its format, loads the workbook via Aspose.Cells LoadOptions, clears workbook‑level encryption, optionally unprotects protected sheets, and writes a decrypted copy with a "_decrypted" suffix. Includes robust error handling and logging for missing files and format detection.
class DecryptBatch
{
    static void Main()
    {
        // Map of encrypted file paths to their corresponding passwords
        var filePasswordMap = new Dictionary<string, string>
        {
            // Example entries – replace with actual file paths and passwords
            { @"C:\EncryptedFiles\Report1.xlsx", "Password123" },
            { @"C:\EncryptedFiles\Report2.xls",  "Secret!@#" }
        };

        foreach (var entry in filePasswordMap)
        {
            string inputPath = entry.Key;
            string password = entry.Value;

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                continue;
            }

            try
            {
                // Detect file format and encryption status (optional, for logging)
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath, password);
                Console.WriteLine($"{Path.GetFileName(inputPath)} – Encrypted: {formatInfo.IsEncrypted}");

                // Load the workbook using the provided password
                var loadOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                var workbook = new Workbook(inputPath, loadOptions);

                // Remove workbook-level encryption before saving
                workbook.Settings.Password = null;

                // If the workbook is protected (not just encrypted), attempt to unprotect it
                try
                {
                    workbook.Unprotect(password);
                }
                catch
                {
                    // Ignore if unprotect fails – the workbook may not be protected
                }

                // Build output file name (original name with "_decrypted" suffix)
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(inputPath) ?? string.Empty,
                    Path.GetFileNameWithoutExtension(inputPath) + "_decrypted" + Path.GetExtension(inputPath));

                // Save the decrypted workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Decrypted file saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{inputPath}': {ex.Message}");
            }
        }
    }
}
