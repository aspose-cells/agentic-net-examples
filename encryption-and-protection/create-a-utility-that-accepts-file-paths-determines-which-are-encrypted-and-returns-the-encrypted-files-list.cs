// Title: C# Utility to Detect Encrypted Excel Files Using Aspose.Cells
// Description: A C# helper that scans a list of file paths, uses Aspose.Cells FileFormatUtil to identify encrypted workbooks via the IsEncrypted flag, logs missing or unsupported files, and returns only the encrypted file paths.
// Keywords: Aspose.Cells encryption detection | C# detect password protected Excel | FileFormatUtil IsEncrypted | list encrypted .xlsx files | bulk Excel security check .NET
// Common Searches: how to find password protected Excel files in C# | Aspose.Cells detect encrypted workbook programmatically | C# filter encrypted spreadsheets before import | list encrypted Excel files using Aspose.Cells
// Developer Intent: Find and collect the paths of Excel workbooks that are encrypted within a given collection.
// Use Cases: Exclude encrypted workbooks from a mass import routine. | Create an audit report of password‑protected spreadsheets on a file server. | Validate uploads in a web portal and reject encrypted Excel files.
// AI Prompts: Generate a C# method that returns encrypted Excel file paths using Aspose.Cells FileFormatUtil. | Extend the EncryptionDetector to also return the encryption algorithm name for each file. | Add retry logic and detailed logging for unsupported formats while detecting encryption.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // A C# helper that scans a list of file paths, uses Aspose.Cells FileFormatUtil to identify encrypted workbooks via the IsEncrypted flag, logs missing or unsupported files, and returns only the encrypted file paths.
    public static class EncryptionDetector
    {
        /// <summary>
        /// Checks each file path and returns a list of files that are encrypted.
        /// </summary>
        /// <param name="filePaths">Collection of file paths to examine.</param>
        /// <returns>List of encrypted file paths.</returns>
        public static List<string> GetEncryptedFiles(IEnumerable<string> filePaths)
        {
            var encryptedFiles = new List<string>();

            foreach (var path in filePaths)
            {
                // Ensure the file exists before attempting detection
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    continue;
                }

                try
                {
                    // Detect file format and retrieve encryption information
                    FileFormatInfo info = FileFormatUtil.DetectFileFormat(path);

                    // If the file is encrypted, add it to the result list
                    if (info.IsEncrypted)
                    {
                        encryptedFiles.Add(path);
                    }
                }
                catch (Exception ex)
                {
                    // Log any errors (e.g., unsupported format) and continue processing other files
                    Console.WriteLine($"Error processing '{path}': {ex.Message}");
                }
            }

            return encryptedFiles;
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Define a set of file paths to check
            var filesToCheck = new List<string>
            {
                "sample1.xlsx",
                "sample2.xlsx",
                "encrypted1.xlsx",
                "encrypted2.xlsm"
            };

            // Get the list of encrypted files
            List<string> encrypted = EncryptionDetector.GetEncryptedFiles(filesToCheck);

            // Output the results
            Console.WriteLine("Encrypted files:");
            foreach (var encFile in encrypted)
            {
                Console.WriteLine(encFile);
            }
        }
    }
}
