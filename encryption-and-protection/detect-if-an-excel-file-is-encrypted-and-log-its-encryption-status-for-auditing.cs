// Title: Detect encrypted Excel workbook and log its encryption status with Aspose.Cells for .NET (C#)
// Common Searches: C# code to check if an Excel file is password protected using Aspose.Cells | how to log encryption status of an Excel workbook in .NET with Aspose.Cells | Aspose.Cells detect encrypted workbook without providing a password | audit Excel file encryption by catching CellsException in C#
// Tags: encrypted workbook detection Aspose.Cells | encryption status logging C# Aspose.Cells | password‑protected Excel handling Aspose.Cells | audit Excel encryption Aspose.Cells | CellsException handling for encrypted files

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that receives an Excel file path, verifies its existence, and attempts to open it with Aspose.Cells. Successful load indicates the workbook is not encrypted; a CellsException containing a password‑related message signals that the file is encrypted and requires a password. The utility logs the encryption status for auditing and captures any other errors.
class Program
{
    static void Main(string[] args)
    {
        // Expect the Excel file path as the first command‑line argument
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program <excel-file-path>");
            return;
        }

        string filePath = args[0];
        LogEncryptionStatus(filePath);
    }

    static void LogEncryptionStatus(string filePath)
    {
        // Prevent FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"[Audit] File not found: \"{filePath}\"");
            return;
        }

        try
        {
            // Attempt to load the workbook without a password.
            // If the file is encrypted, Aspose.Cells throws a CellsException.
            var workbook = new Workbook(filePath);

            // Loading succeeded – the file is not encrypted.
            Console.WriteLine($"[Audit] The file \"{filePath}\" is NOT encrypted.");
        }
        catch (CellsException ex) when (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            // The file is encrypted and requires a password.
            Console.WriteLine($"[Audit] The file \"{filePath}\" is encrypted (password required).");
        }
        catch (CellsException ex)
        {
            // Other Aspose.Cells‑specific errors.
            Console.WriteLine($"[Audit] Failed to open \"{filePath}\": {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected errors.
            Console.WriteLine($"[Audit] Unexpected error opening \"{filePath}\": {ex.Message}");
        }
    }
}
