using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

namespace EncryptTimingDemo
{
    class Program
    {
        // Adjust these paths as needed
        private const string InputFolder = @"C:\Data\LargeDataset";
        private const string OutputFolder = @"C:\Data\EncryptedOutput";
        private const string LogFilePath = @"C:\Data\EncryptionLog.csv";

        // Encryption settings
        private const string Password = "StrongPassword123";
        private const EncryptionType EncType = EncryptionType.StrongCryptographicProvider;
        private const int KeyLength = 128; // 40, 128, or 256

        static void Main()
        {
            try
            {
                // Verify input folder exists
                if (!Directory.Exists(InputFolder))
                {
                    Console.WriteLine($"Input folder not found: {InputFolder}");
                    return;
                }

                // Ensure output directory exists
                Directory.CreateDirectory(OutputFolder);

                // Prepare log file
                using (var logWriter = new StreamWriter(LogFilePath, false))
                {
                    logWriter.WriteLine("FileName,IsInitiallyEncrypted,EncryptionTimeMs,OutputFile");

                    foreach (string filePath in Directory.GetFiles(InputFolder))
                    {
                        // Guard against missing files
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found, skipping: {filePath}");
                            continue;
                        }

                        string fileName = Path.GetFileName(filePath);
                        string outputPath = Path.Combine(OutputFolder, fileName);

                        // Detect if the source file is already encrypted
                        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                        bool initiallyEncrypted = formatInfo.IsEncrypted;

                        // Load workbook (provide password if needed)
                        var loadOptions = new LoadOptions();
                        if (initiallyEncrypted)
                        {
                            // If you know the password for already encrypted files, set it here
                            // loadOptions.Password = "ExistingPassword";
                        }

                        Workbook workbook;
                        try
                        {
                            workbook = new Workbook(filePath, loadOptions);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to load '{fileName}': {ex.Message}");
                            continue;
                        }

                        // Apply encryption settings
                        workbook.Settings.Password = Password;
                        workbook.SetEncryptionOptions(EncType, KeyLength);

                        // Measure encryption (save) time
                        Stopwatch sw = Stopwatch.StartNew();
                        try
                        {
                            workbook.Save(outputPath);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to save encrypted '{fileName}': {ex.Message}");
                            continue;
                        }
                        sw.Stop();

                        // Log results
                        Console.WriteLine($"{fileName} encrypted in {sw.ElapsedMilliseconds} ms");
                        logWriter.WriteLine($"{fileName},{initiallyEncrypted},{sw.ElapsedMilliseconds},{outputPath}");
                    }
                }

                Console.WriteLine("Encryption process completed. Log saved to: " + LogFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}