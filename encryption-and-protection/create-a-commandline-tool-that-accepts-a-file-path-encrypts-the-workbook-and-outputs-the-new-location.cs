// Title: C# Console Tool to Encrypt an Excel Workbook with Aspose.Cells – Command‑Line Utility
// Description: A .NET console application that accepts an Excel file path, optionally a password, loads the workbook using Aspose.Cells, applies password protection, creates an output file with an "_encrypted" suffix, saves the protected workbook, and prints the new location or any error messages. Ideal for scripting, CI/CD, or end‑user protection without opening Excel.
// Keywords: Aspose.Cells | C# console | Excel encryption | workbook password protection | command line utility | CLI Excel encrypt | .NET encrypt workbook | batch Excel security | CI/CD Excel protection | GitHub C# Excel encrypt
// Common Searches: encrypt Excel file using Aspose.Cells C# console app | command line tool to add password to .xlsx | save encrypted workbook with _encrypted suffix .NET | C# CLI encrypt workbook Aspose.Cells example | batch encrypt Excel files in Windows
// Developer Intent: Protect an existing Excel workbook from the command line by applying a password and writing the encrypted file to a new location.
// Use Cases: Automate workbook encryption in a nightly batch script for multiple reports. | Integrate the tool into a CI/CD pipeline to secure generated spreadsheets before deployment. | Provide non‑technical users a simple executable to password‑protect spreadsheets without Microsoft Excel.
// AI Prompts: Write a PowerShell script that runs WorkbookEncryptor.exe for every .xlsx file in a given folder and logs the output paths. | Extend the program to accept an output‑directory argument while preserving the original filename and adding the _encrypted suffix. | Add a command‑line option to choose between AES‑128 and AES‑256 encryption modes supported by Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptor
{
    // A .NET console application that accepts an Excel file path, optionally a password, loads the workbook using Aspose.Cells, applies password protection, creates an output file with an "_encrypted" suffix, saves the protected workbook, and prints the new location or any error messages. Ideal for scripting, CI/CD, or end‑user protection without opening Excel.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a file path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WorkbookEncryptor <inputFilePath> [password]");
                return;
            }

            string inputPath = args[0];

            // Optional password argument; use default if not supplied
            string password = args.Length > 1 ? args[1] : "password123";

            // Validate input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found - {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Set the encryption password
                workbook.Settings.Password = password;

                // Build output file path (insert "_encrypted" before extension)
                string directory = Path.GetDirectoryName(inputPath);
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                string extension = Path.GetExtension(inputPath);
                string outputPath = Path.Combine(directory, $"{filenameWithoutExt}_encrypted{extension}");

                // Save the encrypted workbook
                workbook.Save(outputPath);

                // Inform the user of the new location
                Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
