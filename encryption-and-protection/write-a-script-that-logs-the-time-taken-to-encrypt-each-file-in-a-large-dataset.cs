// Title: C# script to batch encrypt Excel files with Aspose.Cells and log encryption time
// Description: Scans a folder for .xlsx workbooks, skips already encrypted files, applies a password and strong 128‑bit encryption using Aspose.Cells, saves the protected copies, and records the elapsed milliseconds for each file with a Stopwatch.
// Keywords: Aspose.Cells | C# encryption | Excel workbook encryption | batch encrypt Excel | measure encryption time | Stopwatch .NET | strong 128-bit encryption | detect encrypted workbook | file format detection | performance logging
// Common Searches: how to encrypt multiple Excel files with Aspose.Cells C# | measure time taken to encrypt Excel workbook in .NET | skip already encrypted Excel files using Aspose.Cells | log encryption duration for batch Excel processing | set 128-bit password protection for Excel with Aspose.Cells
// Developer Intent: I need to encrypt a collection of Excel workbooks with Aspose.Cells, apply a password, and capture how long each encryption operation takes.
// Use Cases: Processing nightly financial statement batches while monitoring encryption throughput | Compliance audits that require proof of encryption time per document | Optimizing server resources by measuring encryption performance across large datasets | Automating secure archiving of Excel reports with per‑file performance metrics
// AI Prompts: Create a C# example that encrypts .xlsx, .xls, and .xlsm files with Aspose.Cells, records each file’s encryption time, and writes a summary CSV. | Enhance the script to run in parallel threads and output total and average encryption times. | Provide a PowerShell wrapper to invoke the C# encryption timer and store results in Azure Table storage. | Write unit tests that verify password protection, encryption strength, and timing accuracy for various file sizes.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionTimer
{
    // Scans a folder for .xlsx workbooks, skips already encrypted files, applies a password and strong 128‑bit encryption using Aspose.Cells, saves the protected copies, and records the elapsed milliseconds for each file with a Stopwatch.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source workbooks
            string inputFolder = @"C:\Data\Input";
            // Folder where encrypted workbooks will be saved
            string outputFolder = @"C:\Data\Output";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify input folder exists; if not, inform the user and exit gracefully
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                Console.WriteLine("Please create the folder and add Excel files before running the program.");
                return;
            }

            // Password to protect the workbooks
            string password = "SecurePassword123";

            // Iterate over all Excel files in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                try
                {
                    // Detect if the file is already encrypted
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                    if (formatInfo.IsEncrypted)
                    {
                        Console.WriteLine($"{fileName} is already encrypted. Skipping.");
                        continue;
                    }

                    // Ensure the source file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}. Skipping.");
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
                    Console.WriteLine($"{fileName} encrypted in {sw.ElapsedMilliseconds} ms.");
                }
                catch (Exception ex)
                {
                    // Log any errors for the current file and continue with the next one
                    Console.WriteLine($"Error processing {fileName}: {ex.Message}");
                }
            }

            Console.WriteLine("Encryption process completed.");
        }
    }
}
