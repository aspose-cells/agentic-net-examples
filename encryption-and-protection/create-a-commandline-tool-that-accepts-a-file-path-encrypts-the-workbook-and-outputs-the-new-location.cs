// Title: C# CLI tool to encrypt an Excel workbook with Aspose.Cells and save as a new file
// Description: A .NET console application that accepts an Excel file path and an optional password from the command line, loads the workbook using Aspose.Cells, applies password protection (default or supplied), optionally sets AES encryption, creates an output name with “_encrypted”, and saves the protected workbook. Includes file‑existence checks and exception handling.
// Keywords: Aspose.Cells | C# encrypt Excel | CLI Excel encryption | password protect workbook | encrypt .xlsx command line | set workbook password | AES 128 Excel encryption | Aspose.Cells SetEncryptionOptions | batch encrypt Excel files | console app encrypt workbook
// Common Searches: encrypt excel file using Aspose.Cells C# | command line tool to password protect xlsx | how to set workbook password in .NET | save encrypted workbook with Aspose.Cells | CLI encrypt Excel workbook Aspose | default password for encrypted workbook C#
// Developer Intent: Encrypt an existing Excel file from the command line and output a password‑protected copy.
// Use Cases: Automate encryption of generated reports in CI/CD pipelines. | Run as a scheduled task to secure archived spreadsheets. | Integrate into batch scripts for bulk workbook protection. | Provide a fallback password when none is supplied to guarantee security.
// AI Prompts: Write a PowerShell script that loops through all .xlsx files in a folder and calls WorkbookEncryptor.exe with a specified password. | Extend the program to accept an output directory argument and create the encrypted file there. | Add a command‑line switch to choose AES‑128 or AES‑256 encryption via Aspose.Cells SetEncryptionOptions. | Implement logging to a file for each encryption operation. | Create a Dockerfile to run the CLI tool in a container.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptor
{
    // A .NET console application that accepts an Excel file path and an optional password from the command line, loads the workbook using Aspose.Cells, applies password protection (default or supplied), optionally sets AES encryption, creates an output name with “_encrypted”, and saves the protected workbook. Includes file‑existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect at least the input file path; optional second argument is the password.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WorkbookEncryptor <inputFilePath> [password]");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found - {inputPath}");
                return;
            }

            // Use the second argument as password if provided, otherwise use a default password.
            string password = args.Length > 1 ? args[1] : "defaultPassword";

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Set the encryption password.
                workbook.Settings.Password = password;

                // Optionally, you can set stronger encryption options (e.g., AES 128-bit).
                // workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // Build the output file name by inserting "_encrypted" before the extension.
                string directory = Path.GetDirectoryName(inputPath);
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                string extension = Path.GetExtension(inputPath);
                string outputPath = Path.Combine(directory, $"{filenameWithoutExt}_encrypted{extension}");

                // Save the encrypted workbook.
                workbook.Save(outputPath);

                Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during encryption: {ex.Message}");
            }
        }
    }
}
