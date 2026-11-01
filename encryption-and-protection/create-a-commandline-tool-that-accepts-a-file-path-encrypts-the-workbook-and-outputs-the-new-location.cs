// Title: C# command‑line tool to encrypt an Excel workbook with Aspose.Cells and save it with an _encrypted suffix
// Common Searches: how to encrypt an existing .xlsx file using Aspose.Cells in a C# console application | C# command line program to add password protection to an Excel workbook with Aspose.Cells | save encrypted Excel workbook to a new file name using Aspose.Cells CLI tool | encrypt Excel workbook from file path argument and output encrypted file location in C# | Aspose.Cells encrypt workbook from command line and specify output directory
// Tags: Aspose.Cells workbook password encryption | C# console tool for Excel file protection | command‑line generation of encrypted .xlsx | output path with _encrypted suffix for Excel | Workbook.Settings.Password usage in .NET

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that takes an input file path, loads the workbook with Aspose.Cells, sets a static password via Workbook.Settings.Password to encrypt the file, builds an output path by appending "_encrypted" to the original name, saves the encrypted workbook as XLSX, and writes the new file location to the console while handling missing arguments, missing files, and runtime exceptions.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Verify that a file path argument was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: encryptworkbook <inputFilePath>");
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
            workbook.Settings.Password = "Password123"; // Replace or obtain from user input as needed

            // Build the output file path (original name with "_encrypted" suffix)
            string directory = Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
            string outputPath = Path.Combine(directory, $"{fileNameWithoutExt}_encrypted.xlsx");

            // Save the encrypted workbook to the new location
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Output the location of the encrypted file
            Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
