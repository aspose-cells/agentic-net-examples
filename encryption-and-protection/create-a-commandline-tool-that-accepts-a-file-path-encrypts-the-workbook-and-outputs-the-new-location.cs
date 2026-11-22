// Title: Create a C# command‑line tool that encrypts an Excel workbook with a password using Aspose.Cells and saves it with an _encrypted suffix
// Common Searches: how to add password protection to an Excel file using Aspose.Cells in a C# console app | C# command line program to encrypt a .xlsx workbook and output to a new file | Aspose.Cells encrypt workbook from file path argument and save with custom suffix | encrypt Excel workbook via command line and get encrypted file path in C# | save password‑protected Excel workbook using Aspose.Cells console application
// Tags: Aspose.Cells workbook password encryption | C# console utility for Excel file protection | Workbook.Settings.Password API usage | command‑line Excel encryption with output suffix | generate encrypted file path programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptor
{
    // A C# console application that accepts an Excel file path, loads the workbook with Aspose.Cells, applies a password via Workbook.Settings.Password, builds a new filename by appending "_encrypted", saves the protected workbook, and prints the encrypted file location, with basic argument validation and error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a file path argument was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WorkbookEncryptor <input-file-path>");
                return;
            }

            string inputPath = args[0];

            // Validate input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found - {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified path
                Workbook workbook = new Workbook(inputPath);

                // Set a password to encrypt the workbook.
                // Change the password value as needed.
                string password = "MySecretPassword";
                workbook.Settings.Password = password;

                // Build the output file path (same folder, with "_encrypted" suffix)
                string directory = Path.GetDirectoryName(inputPath);
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                string extension = Path.GetExtension(inputPath);
                string outputPath = Path.Combine(directory, $"{filenameWithoutExt}_encrypted{extension}");

                // Save the encrypted workbook to the new location
                workbook.Save(outputPath);

                // Output the new location
                Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
