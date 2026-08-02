// Title: C# Batch Encrypt Excel Files with Aspose.Cells and Record Per‑File Timing
// Description: A console utility that scans a source directory for .xlsx workbooks, skips already‑encrypted files, applies a password with 128‑bit strong encryption using Aspose.Cells, saves the protected copies to a target folder, and logs the elapsed time for each encryption operation.
// Keywords: Aspose.Cells C# encryption | batch encrypt Excel workbooks | measure encryption performance .NET | password protect .xlsx files | stopwatch timing Aspose.Cells | large dataset Excel security
// Common Searches: how to encrypt multiple Excel files with Aspose.Cells in C# | C# script to log encryption time for each workbook | batch password protect .xlsx using Aspose.Cells | measure Aspose.Cells encryption speed | encrypt Excel files in bulk and track duration
// Developer Intent: Automatically protect a collection of Excel spreadsheets with a password using Aspose.Cells while capturing the processing time for each file.
// Use Cases: Compliance‑driven archiving of thousands of spreadsheets with performance metrics | Data‑migration pipelines that need to skip already‑encrypted workbooks | Generating logs for monitoring encryption throughput in on‑prem or cloud environments | Creating benchmark reports for Aspose.Cells encryption capabilities
// AI Prompts: Generate a C# method that accepts input and output folder paths and returns a dictionary of workbook names and their encryption times in milliseconds using Aspose.Cells. | Show how to modify the program to write file names and elapsed times to a CSV file instead of console output. | Provide code to parallelize the encryption loop safely while preserving accurate per‑file timing logs.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionTimer
{
    // A console utility that scans a source directory for .xlsx workbooks, skips already‑encrypted files, applies a password with 128‑bit strong encryption using Aspose.Cells, saves the protected copies to a target folder, and logs the elapsed time for each encryption operation.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing the original workbooks
            string inputFolder = @"C:\Data\Workbooks";
            // Output folder where encrypted workbooks will be saved
            string outputFolder = @"C:\Data\EncryptedWorkbooks";

            try
            {
                // Verify input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Password to use for encryption
                const string password = "SecurePassword123";

                // Process each Excel file in the input folder (including subfolders)
                foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories))
                {
                    try
                    {
                        string fileName = Path.GetFileName(filePath);
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Detect if the file is already encrypted
                        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                        if (formatInfo.IsEncrypted)
                        {
                            Console.WriteLine($"{fileName} is already encrypted. Skipping.");
                            continue;
                        }

                        // Verify the file exists before loading
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Start timing the encryption process
                        Stopwatch sw = Stopwatch.StartNew();

                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Set password protection
                        workbook.Settings.Password = password;

                        // Set encryption options (strong encryption with 128‑bit key)
                        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                        // Save the encrypted workbook
                        workbook.Save(outputPath);

                        // Stop timing
                        sw.Stop();

                        // Log the time taken
                        Console.WriteLine($"{fileName}: Encryption completed in {sw.ElapsedMilliseconds} ms");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Encryption process completed for all files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
