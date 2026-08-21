// Title: Batch encrypt XLS workbooks with a single password using Aspose.Cells for .NET
// Description: Iterates over all *.xls files in a source folder, loads each workbook with Aspose.Cells, applies a common password (optionally setting encryption strength), and saves the protected copies to a target directory while handling missing files and runtime errors.
// Keywords: Aspose.Cells C# encrypt XLS | batch Excel password protection | programmatic Excel encryption .NET | set workbook password Aspose | encrypt multiple Excel files
// Common Searches: C# batch encrypt XLS files Aspose.Cells | apply same password to many Excel workbooks | automate Excel file encryption .NET | Aspose.Cells encrypt all files in folder | set encryption options for XLS with Aspose
// Developer Intent: Apply one password to every .xls workbook in a folder and write the encrypted versions to another location.
// Use Cases: Secure a collection of legacy financial reports before archiving on a shared drive. | Enforce company‑wide password protection on exported spreadsheets generated nightly. | Integrate automatic encryption of incoming XLS files in a data‑processing pipeline to satisfy compliance policies.
// AI Prompts: Write C# code that encrypts all .xls files in a specified directory with Aspose.Cells, allowing the password and encryption algorithm to be passed as parameters. | Show how to use Aspose.Cells SetEncryptionOptions to apply AES‑256 encryption to a batch of workbooks. | Provide best‑practice guidelines for logging and error handling when encrypting multiple Excel files programmatically.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchEncryptXls
{
    // Iterates over all *.xls files in a source folder, loads each workbook with Aspose.Cells, applies a common password (optionally setting encryption strength), and saves the protected copies to a target directory while handling missing files and runtime errors.
    class Program
    {
        static void Main()
        {
            // Source folder containing the original XLS files
            string sourceFolder = @"C:\SourceFolder";

            // Target folder where encrypted files will be saved
            string targetFolder = @"C:\TargetFolder";

            // Common password to protect all workbooks
            const string password = "CommonPassword123";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                // Ensure the target directory exists
                Directory.CreateDirectory(targetFolder);

                // Process each .xls file in the source folder
                foreach (string sourceFilePath in Directory.GetFiles(sourceFolder, "*.xls"))
                {
                    try
                    {
                        // Verify the source file exists
                        if (!File.Exists(sourceFilePath))
                        {
                            Console.WriteLine($"File not found: {sourceFilePath}");
                            continue;
                        }

                        // Load the workbook from the source file
                        Workbook workbook = new Workbook(sourceFilePath);

                        // Set the password that will encrypt the workbook
                        workbook.Settings.Password = password;

                        // Optional: specify encryption algorithm and key length
                        // workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                        // Build the destination file path (same file name, different folder)
                        string destFilePath = Path.Combine(targetFolder, Path.GetFileName(sourceFilePath));

                        // Save the encrypted workbook to the target location
                        workbook.Save(destFilePath);
                        Console.WriteLine($"Encrypted: {Path.GetFileName(sourceFilePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{sourceFilePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch encryption completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
