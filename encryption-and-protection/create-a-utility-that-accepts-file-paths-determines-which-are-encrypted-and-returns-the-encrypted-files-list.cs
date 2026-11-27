// Title: C# utility to identify password‑protected Excel workbooks with Aspose.Cells
// Common Searches: how to find encrypted .xlsx files with Aspose.Cells C# | list password protected Excel workbooks in a folder using Aspose.Cells | C# check if an Excel workbook requires a password without opening it | detect workbook encryption during load with Aspose.Cells | skip missing files while scanning Excel files for encryption in C#
// Tags: detect encrypted Excel workbook Aspose.Cells | identify password protected .xlsx C# | enumerate encrypted files using Aspose.Cells | ignore missing files when scanning Excel | use CellsException to determine workbook encryption

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The provided utility iterates over a collection of file paths, ignores any that do not exist, and attempts to load each workbook with Aspose.Cells. If loading throws a CellsException, the file is considered encrypted (password‑protected) and its path is added to the result list, which is then returned.
public static class ExcelEncryptionUtility
{
    public static List<string> GetEncryptedFiles(IEnumerable<string> filePaths)
    {
        var encryptedFiles = new List<string>();

        foreach (var path in filePaths)
        {
            // Skip non‑existent files
            if (!File.Exists(path))
                continue;

            try
            {
                // Attempt to load the workbook without a password.
                // If the file is encrypted, Aspose.Cells throws a CellsException.
                using (var workbook = new Workbook(path))
                {
                    // Loaded successfully – not encrypted.
                }
            }
            catch (CellsException)
            {
                // Treat any CellsException as indication of encryption.
                encryptedFiles.Add(path);
            }
            catch (Exception)
            {
                // Non‑Aspose exceptions (e.g., IO errors) are ignored for this utility.
            }
        }

        return encryptedFiles;
    }
}

// Simple entry point for demonstration/testing.
public class Program
{
    public static void Main(string[] args)
    {
        // Example file list; replace with actual paths as needed.
        var files = new List<string>
        {
            "sample1.xlsx",
            "sample2.xlsx"
        };

        try
        {
            var encrypted = ExcelEncryptionUtility.GetEncryptedFiles(files);
            Console.WriteLine("Encrypted files:");
            foreach (var f in encrypted)
            {
                Console.WriteLine(f);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
