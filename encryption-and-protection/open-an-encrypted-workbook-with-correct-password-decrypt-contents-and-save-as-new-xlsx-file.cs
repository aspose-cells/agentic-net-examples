// Title: Decrypt a password‑protected XLSX workbook and save it as an unencrypted file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to open a password‑protected Excel file and save it without encryption | How to remove password protection from an XLSX workbook using Aspose.Cells .NET | Load an encrypted XLSX with LoadOptions and export an unencrypted copy in C# | Example of decrypting a protected Excel workbook with Aspose.Cells and SaveFormat.Xlsx | Programmatically clear workbook password in Aspose.Cells for .NET
// Tags: encrypted workbook loading LoadOptions Aspose.Cells | clear workbook password Aspose.Cells .NET | save workbook unencrypted SaveFormat.Xlsx | Aspose.Cells decrypt Excel file C# | remove Excel encryption Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// The sample checks that the source encrypted.xlsx exists, loads it with a LoadOptions object that supplies the correct password, clears the workbook's password property to strip encryption, and then saves the result as decrypted.xlsx in XLSX format, handling both Aspose.Cells‑specific and general exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string sourcePath = "encrypted.xlsx";

        // Password used to encrypt the workbook
        string password = "myPassword";

        // Path for the decrypted output workbook
        string destPath = "decrypted.xlsx";

        try
        {
            // Ensure the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the encrypted workbook with the correct password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Remove encryption by clearing the password
            workbook.Settings.Password = string.Empty;

            // Save the workbook as a new XLSX file without encryption
            workbook.Save(destPath, SaveFormat.Xlsx);
            Console.WriteLine($"Decrypted workbook saved to: {destPath}");
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
