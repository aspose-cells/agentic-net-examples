// Title: Batch decrypt multiple password‑protected XLSX workbooks with Aspose.Cells for .NET and record results
// Common Searches: C# Aspose.Cells batch decrypt password‑protected Excel files in a folder | How to try a list of passwords when opening encrypted XLSX using Aspose.Cells | Log successful and failed decryption attempts while processing many workbooks in .NET | Remove passwords from a collection of XLSX files automatically with Aspose.Cells | Save decrypted Excel workbook without password using Aspose.Cells LoadOptions
// Tags: mass password removal from XLSX Aspose.Cells | load workbook with multiple passwords .NET | save workbook without encryption Aspose.Cells | decryption operation logging C# | process encrypted Excel folder Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The example iterates over all *.xlsx files in a specified input directory, attempts to open each encrypted workbook using a predefined list of passwords via Aspose.Cells LoadOptions, saves the workbook without a password to an output directory when successful, and writes a detailed log entry for each success or failure.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing encrypted XLSX files
            string inputFolder = @"C:\EncryptedFiles";
            // Folder to save decrypted files
            string outputFolder = @"C:\DecryptedFiles";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // List of possible passwords
            List<string> passwords = new List<string> { "password1", "password2", "123456", "secret" };

            // Log file for successes and failures
            string logPath = Path.Combine(outputFolder, "decryption_log.txt");
            using (StreamWriter logWriter = new StreamWriter(logPath, false))
            {
                // Iterate over each XLSX file in the input folder
                foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    if (!File.Exists(filePath))
                    {
                        logWriter.WriteLine($"{DateTime.Now}: SKIPPED - File not found \"{filePath}\".");
                        continue;
                    }

                    bool decrypted = false;

                    // Try each password until one works
                    foreach (string pwd in passwords)
                    {
                        try
                        {
                            // Load the workbook using the current password
                            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                            {
                                Password = pwd
                            };
                            Workbook wb = new Workbook(filePath, loadOptions);

                            // Save the workbook without a password (decrypted)
                            string fileName = Path.GetFileNameWithoutExtension(filePath);
                            string outputPath = Path.Combine(outputFolder, fileName + "_decrypted.xlsx");
                            wb.Save(outputPath, SaveFormat.Xlsx);

                            // Log success and stop trying other passwords for this file
                            logWriter.WriteLine($"{DateTime.Now}: SUCCESS - \"{filePath}\" decrypted with password \"{pwd}\".");
                            decrypted = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            // Loading failed – likely due to incorrect password; continue with next password
                            // Optionally log detailed error for debugging
                            // logWriter.WriteLine($"{DateTime.Now}: INFO - Attempt with password \"{pwd}\" failed: {ex.Message}");
                        }
                    }

                    if (!decrypted)
                    {
                        // Log failure after all passwords have been tried
                        logWriter.WriteLine($"{DateTime.Now}: FAILURE - \"{filePath}\" could not be decrypted with provided passwords.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Global exception handling
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
