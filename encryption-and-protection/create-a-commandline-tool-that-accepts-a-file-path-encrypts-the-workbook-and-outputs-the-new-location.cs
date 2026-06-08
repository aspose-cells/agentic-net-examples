using System;
using Aspose.Cells;

namespace WorkbookEncryptor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: WorkbookEncryptor <inputFilePath> [password] [outputFilePath]");
                return;
            }

            string inputPath = args[0];
            string password = args.Length >= 2 ? args[1] : "defaultPassword";

            // Determine output path
            string outputPath;
            if (args.Length >= 3)
            {
                outputPath = args[2];
            }
            else
            {
                // Append "_encrypted" before the extension
                string extension = System.IO.Path.GetExtension(inputPath);
                string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(inputPath);
                string directory = System.IO.Path.GetDirectoryName(inputPath);
                outputPath = System.IO.Path.Combine(directory, $"{fileNameWithoutExt}_encrypted{extension}");
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set the password to encrypt the workbook
            workbook.Settings.Password = password;

            // Save the encrypted workbook to the new location
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook encrypted and saved to: {outputPath}");
        }
    }
}