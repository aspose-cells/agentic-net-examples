// Title: C# Utility to Detect Encrypted Excel Workbooks with Aspose.Cells FileFormatUtil
// Description: A C# helper that iterates over supplied file paths, confirms each file exists, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to read the IsEncrypted flag, and returns a list of only the encrypted workbook paths while gracefully handling missing files and exceptions.
// Keywords: Aspose.Cells | FileFormatUtil | IsEncrypted | C# encrypted Excel detection | password protected workbook | detect encrypted spreadsheet .NET | Excel file encryption check | bulk encrypted file scan | Aspose.Cells API | detect encrypted files
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells C# | C# list of password‑protected Excel workbooks | detect encrypted spreadsheets in a folder with Aspose.Cells | FileFormatUtil IsEncrypted example | filter encrypted Excel files before conversion .NET
// Developer Intent: Identify which Excel workbooks in a given collection are encrypted and retrieve their file paths.
// Use Cases: Skip password‑protected spreadsheets during bulk import or conversion. | Create compliance reports that list encrypted workbooks in a directory. | Validate a batch of files before applying automated data extraction. | Log encrypted files for security audits in enterprise environments.
// AI Prompts: Generate a C# method that accepts an IEnumerable<string> of file paths and returns only those marked as encrypted by Aspose.Cells. | Refactor the GetEncryptedFiles function to write missing‑file and error messages to a structured log file instead of the console. | Extend the utility to output the encryption algorithm (if available) alongside each encrypted workbook path.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// A C# helper that iterates over supplied file paths, confirms each file exists, uses Aspose.Cells.FileFormatUtil.DetectFileFormat to read the IsEncrypted flag, and returns a list of only the encrypted workbook paths while gracefully handling missing files and exceptions.
public class EncryptionDetector
{
    // Returns a list of file paths that are encrypted.
    public static List<string> GetEncryptedFiles(IEnumerable<string> filePaths)
    {
        var encryptedFiles = new List<string>();

        foreach (var path in filePaths)
        {
            try
            {
                // Verify the file exists before attempting detection.
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    continue;
                }

                // Detect file format information for the given file.
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(path);

                // If the file is encrypted, add it to the result list.
                if (info.IsEncrypted)
                {
                    encryptedFiles.Add(path);
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors and continue processing other files.
                Console.WriteLine($"Error processing '{path}': {ex.Message}");
            }
        }

        return encryptedFiles;
    }

    // Example entry point demonstrating usage.
    public static void Main()
    {
        var filesToCheck = new List<string>
        {
            "example.xlsx",
            "encrypted.xlsx",
            "plain.xls"
        };

        List<string> encrypted = GetEncryptedFiles(filesToCheck);

        Console.WriteLine("Encrypted files:");
        foreach (var file in encrypted)
        {
            Console.WriteLine(file);
        }
    }
}
