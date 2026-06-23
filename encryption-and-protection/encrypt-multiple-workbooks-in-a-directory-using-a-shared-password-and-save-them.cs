using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    public class EncryptWorkbooksInDirectory
    {
        /// <summary>
        /// Encrypts all Excel workbooks in the specified source directory with the given password
        /// and saves the encrypted copies to the destination directory.
        /// </summary>
        /// <param name="sourceDirectory">Path to the folder containing original workbooks.</param>
        /// <param name="destinationDirectory">Path to the folder where encrypted workbooks will be saved.</param>
        /// <param name="password">Password to protect each workbook.</param>
        public static void Run(string sourceDirectory, string destinationDirectory, string password)
        {
            // Ensure source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
                return;
            }

            // Ensure the destination directory exists
            if (!Directory.Exists(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }

            // Define the file extensions to process
            string[] extensions = new[] { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb", "*.ods" };

            foreach (string ext in extensions)
            {
                // Get all files with the current extension
                string[] files = Directory.GetFiles(sourceDirectory, ext, SearchOption.TopDirectoryOnly);
                foreach (string filePath in files)
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook from the file
                        using (Workbook workbook = new Workbook(filePath))
                        {
                            // Set the password to encrypt the workbook
                            workbook.Settings.Password = password;

                            // Build the output file path (preserve original file name)
                            string fileName = Path.GetFileName(filePath);
                            string outputPath = Path.Combine(destinationDirectory, fileName);

                            // Save the encrypted workbook (overwrites if file already exists)
                            workbook.Save(outputPath);
                            Console.WriteLine($"Encrypted and saved: {outputPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }
            }
        }
    }

    public class Program
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            // Expecting three arguments: sourceDirectory destinationDirectory password
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: AsposeCellsEncryptionDemo <sourceDirectory> <destinationDirectory> <password>");
                return;
            }

            string sourceDir = args[0];
            string destDir = args[1];
            string pwd = args[2];

            try
            {
                EncryptWorkbooksInDirectory.Run(sourceDir, destDir, pwd);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}