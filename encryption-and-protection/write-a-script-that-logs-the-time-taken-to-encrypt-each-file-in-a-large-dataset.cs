// Title: C# script to encrypt a batch of Excel (.xlsx) files with Aspose.Cells and log per‑file processing time
// Common Searches: how to encrypt multiple .xlsx files using Aspose.Cells in C# and record execution time | C# program to batch encrypt Excel workbooks with password and write timing log | measure encryption duration for each Excel file with Aspose.Cells library | preserve folder structure while encrypting Excel files in C# using Aspose.Cells | log encryption performance of a large Excel dataset to a text file in .NET
// Tags: batch encryption of Excel files with Aspose.Cells | Aspose.Cells password protection for .xlsx | measure encryption time using Stopwatch in C# | log processing duration to text file in .NET | preserve directory hierarchy when saving encrypted workbooks | performance logging for Aspose.Cells encryption

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The C# utility scans a source directory (including subfolders) for .xlsx files, loads each workbook with Aspose.Cells, applies a password to encrypt it, saves the encrypted copy to a target folder while preserving the original folder hierarchy, measures the encryption time with Stopwatch, and writes a timestamped entry containing the relative file path and elapsed milliseconds to EncryptionLog.txt; any errors are captured and logged per file.
class EncryptAndLog
{
    static void Main(string[] args)
    {
        // Input folder containing the original Excel files
        string inputFolder = @"C:\Data\OriginalFiles";
        // Output folder where encrypted files will be saved
        string outputFolder = @"C:\Data\EncryptedFiles";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Log file to record processing times
        string logFilePath = Path.Combine(outputFolder, "EncryptionLog.txt");
        using (StreamWriter logWriter = new StreamWriter(logFilePath, false))
        {
            // Iterate over all Excel files in the input folder (including subfolders)
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories))
            {
                // Skip if the file somehow does not exist
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                try
                {
                    // Measure time for each file
                    Stopwatch sw = Stopwatch.StartNew();

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Set encryption password
                    workbook.Settings.Password = "StrongPassword123";

                    // Determine output file path preserving relative structure
                    string relativePath = Path.GetRelativePath(inputFolder, filePath);
                    string encryptedFilePath = Path.Combine(outputFolder, relativePath);
                    Directory.CreateDirectory(Path.GetDirectoryName(encryptedFilePath));

                    // Save the encrypted workbook
                    workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

                    // Stop timing
                    sw.Stop();

                    // Log the result
                    string logEntry = $"{DateTime.Now:u} | File: {relativePath} | TimeTaken(ms): {sw.ElapsedMilliseconds}";
                    Console.WriteLine(logEntry);
                    logWriter.WriteLine(logEntry);
                }
                catch (Exception ex)
                {
                    // Log any errors for this file and continue processing others
                    string errorEntry = $"{DateTime.Now:u} | File: {filePath} | Error: {ex.Message}";
                    Console.WriteLine(errorEntry);
                    logWriter.WriteLine(errorEntry);
                }
            }
        }

        Console.WriteLine("Encryption process completed. See log for details.");
    }
}
