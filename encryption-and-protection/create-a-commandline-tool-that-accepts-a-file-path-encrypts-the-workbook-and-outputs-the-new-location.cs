// Title: C# CLI Tool to Encrypt an Excel Workbook with Aspose.Cells and Save as _encrypted
// Description: A console application that accepts a file path, validates the file, loads the workbook with Aspose.Cells, sets a password, applies strong 128‑bit encryption, builds an output name with an “_encrypted” suffix, and saves the protected workbook.
// Keywords: Aspose.Cells | C# encryption | Excel password protection | CLI tool | command line | strong cryptographic provider | 128‑bit encryption | Workbook.Save | console application | file path argument
// Common Searches: encrypt Excel file C# Aspose.Cells | command line password protection for .xlsx | C# console app encrypt workbook | Aspose.Cells set encryption options | save encrypted workbook with custom name | batch encrypt Excel files using Aspose
// Developer Intent: Encrypt an existing Excel workbook from the command line and write the encrypted copy to a new location.
// Use Cases: Run the utility in a batch script to protect multiple spreadsheets before archiving. | Integrate the tool into a CI/CD pipeline to secure confidential reports generated during builds. | Schedule a Windows task that automatically encrypts newly created workbooks in a shared folder.
// AI Prompts: Create a C# console program that takes a file path, loads the workbook with Aspose.Cells, applies a password and 128‑bit strong encryption, and saves it with an "_encrypted" suffix. | Add optional command‑line arguments for custom password, encryption type, and key length to the Aspose.Cells encryption CLI tool. | Explain how to extend the program to handle .xls, .xlsx, and .csv files while preserving encryption settings.

using System;
using System.IO;
using Aspose.Cells;

// A console application that accepts a file path, validates the file, loads the workbook with Aspose.Cells, sets a password, applies strong 128‑bit encryption, builds an output name with an “_encrypted” suffix, and saves the protected workbook.
class Program
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: encrypt <inputFilePath>");
            return;
        }

        string inputPath = args[0];

        // Check that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Set a password to encrypt the workbook
        workbook.Settings.Password = "password123";

        // Optionally specify encryption algorithm and key length
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Build the output file path (original name with "_encrypted" suffix)
        string directory = Path.GetDirectoryName(inputPath);
        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
        string extension = Path.GetExtension(inputPath);
        string outputPath = Path.Combine(directory, $"{fileNameWithoutExt}_encrypted{extension}");

        // Save the encrypted workbook
        workbook.Save(outputPath);

        Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
    }
}
