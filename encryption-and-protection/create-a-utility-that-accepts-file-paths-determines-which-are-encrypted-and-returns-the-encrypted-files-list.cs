// Title: C# utility to scan multiple files and list password‑protected Excel workbooks using Aspose.Cells
// AI Prompts: Generate a C# method that takes an IEnumerable<string> of file paths, attempts to load each with Aspose.Cells Workbook, and returns a List<string> of paths where a CellsException indicates a password is required. | Show how to catch Aspose.Cells CellsException and examine its message to detect encrypted .xlsx or .xls files without supplying a password. | Create a console application that reads file paths from command‑line arguments, invokes the encryption‑checking method, and prints the names of all encrypted Excel files.
// Common Searches: how to programmatically detect password protected Excel files in .NET using Aspose.Cells | C# batch scan folder for encrypted .xlsx files without opening them | Aspose.Cells identify workbook encryption by catching CellsException | list encrypted Excel workbooks from a list of file paths in C# | detect password required Excel workbook when loading with Aspose.Cells
// Tags: Aspose.Cells detect encrypted workbook | C# check Excel file password protection | batch enumerate password‑protected xlsx files | handle CellsException for workbook encryption | list encrypted Excel files .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example defines an ExcelEncryptionChecker class with a GetEncryptedFiles method that iterates over supplied file paths, verifies existence, attempts to load each workbook using Aspose.Cells, and captures CellsException containing a password‑related message to identify encrypted Excel files; a console Program demonstrates calling the method with command‑line arguments and printing the encrypted file list.
public static class ExcelEncryptionChecker
{
    // Accepts a collection of file paths and returns the list of encrypted Excel files.
    public static List<string> GetEncryptedFiles(IEnumerable<string> filePaths)
    {
        var encryptedFiles = new List<string>();

        foreach (var filePath in filePaths)
        {
            // Skip if the file does not exist.
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Attempt to load the workbook without providing a password.
                // If the file is encrypted, Aspose.Cells throws a CellsException.
                var workbook = new Workbook(filePath);
                // If loading succeeds, the file is not password‑protected.
            }
            catch (CellsException ex)
            {
                // Check if the exception indicates a password is required.
                // Different Aspose.Cells versions expose this via the message.
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    encryptedFiles.Add(filePath);
                }
                // Other CellsException types are ignored (e.g., corrupted file).
            }
            catch (Exception)
            {
                // Any other exception (e.g., unsupported format) is ignored.
            }
        }

        return encryptedFiles;
    }
}

public class Program
{
    // Entry point required for the console application.
    public static void Main(string[] args)
    {
        try
        {
            // Example usage: provide file paths via command‑line arguments or hard‑code them here.
            IEnumerable<string> filesToCheck = args.Length > 0
                ? args
                : new List<string>
                {
                    // Add sample file paths for testing.
                    "Sample1.xlsx",
                    "Sample2.xls",
                    "EncryptedFile.xlsx"
                };

            var encrypted = ExcelEncryptionChecker.GetEncryptedFiles(filesToCheck);

            Console.WriteLine("Encrypted Excel files found:");
            foreach (var path in encrypted)
            {
                Console.WriteLine(path);
            }

            if (encrypted.Count == 0)
            {
                Console.WriteLine("No encrypted files detected.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions to prevent the program from crashing.
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
