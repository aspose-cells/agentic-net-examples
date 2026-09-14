// Title: Create a C# command‑line tool that encrypts an Excel workbook with a password using Aspose.Cells and saves it to a specified output file
// AI Prompts: Write a C# console program that accepts an input Excel path, an output path, and an optional password, loads the workbook with Aspose.Cells, sets workbook.Settings.Password, and saves the encrypted file. | Implement robust argument validation for the encryption utility, including checks for missing parameters, file existence, and default password fallback. | Extend the tool to allow the user to choose the output format (XLSX or XLS) while keeping the workbook password protection using Aspose.Cells SaveFormat.
// Common Searches: c# aspocells command line encrypt excel workbook password | how to set password on workbook using Aspose.Cells in a console app | encrypt xlsx file from command line with Aspose.Cells .NET | aspocells workbook.Settings.Password example for console utility
// Tags: Aspose.Cells workbook password encryption | C# console Excel file protection | encrypt Excel workbook to XLSX with Aspose.Cells | command‑line workbook.Settings.Password usage | validate input and output paths in .NET file encryption

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptor
{
    // A C# console application that validates command‑line arguments, ensures the input Excel file exists, loads it with Aspose.Cells, applies a password via workbook.Settings.Password, and saves the encrypted workbook to the user‑specified output location.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WorkbookEncryptor <inputFilePath> <outputFilePath> [password]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string password = args.Length >= 3 ? args[2] : "defaultPassword";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Set the password to encrypt the workbook
                workbook.Settings.Password = password;

                // Save the encrypted workbook to the new location
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook encrypted successfully. Saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
