// Title: C# utility to identify password‑protected Excel workbooks using Aspose.Cells
// Common Searches: C# code to find password protected Excel files with Aspose.Cells | How to detect encrypted .xlsx files without supplying a password in .NET | List all password‑protected Excel workbooks in a directory using Aspose.Cells | Determine if an Excel workbook is encrypted programmatically in C# | Batch check multiple Excel files for password protection via Aspose.Cells exception handling
// Tags: Aspose.Cells encrypted workbook detection | C# list password protected .xlsx files | Workbook encryption check using CellsException | Identify protected Excel files .NET | Batch Excel encryption detection Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The ExcelEncryptionChecker class scans a collection of file paths, attempts to load each workbook with Aspose.Cells without a password, and treats any CellsException as an indication of password protection, returning a list of the encrypted file paths.
public static class ExcelEncryptionChecker
{
    /// <param name="filePaths">Collection of file paths to examine.</param>
    /// <returns>List of encrypted file paths.</returns>
    public static List<string> GetEncryptedFiles(IEnumerable<string> filePaths)
    {
        var encryptedFiles = new List<string>();

        foreach (var path in filePaths)
        {
            // Skip non‑existent entries
            if (!File.Exists(path))
                continue;

            try
            {
                // Attempt to load the workbook without providing a password.
                // If the file is encrypted, Aspose.Cells throws a CellsException.
                var workbook = new Workbook(path);
                // Loaded successfully – not encrypted.
            }
            catch (CellsException)
            {
                // Treat any CellsException as indication of password protection.
                encryptedFiles.Add(path);
            }
            catch (Exception)
            {
                // Other exceptions (e.g., invalid format) are ignored for this utility.
            }
        }

        return encryptedFiles;
    }
}

public class Program
{
    /// <summary>
    /// Entry point for the console application.
    /// </summary>
    /// <param name="args">File paths to be checked (optional).</param>
    public static void Main(string[] args)
    {
        try
        {
            // If no arguments are supplied, demonstrate with a sample list.
            var filesToCheck = args.Length > 0
                ? (IEnumerable<string>)args
                : new List<string>
                {
                    // Add sample file paths here for testing, e.g.:
                    // @"C:\Temp\Encrypted.xlsx",
                    // @"C:\Temp\Plain.xlsx"
                };

            var encrypted = ExcelEncryptionChecker.GetEncryptedFiles(filesToCheck);

            Console.WriteLine("Encrypted files found:");
            foreach (var file in encrypted)
            {
                Console.WriteLine(file);
            }

            if (encrypted.Count == 0)
            {
                Console.WriteLine("No encrypted files detected.");
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors.
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
