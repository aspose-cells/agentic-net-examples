// Title: Batch decrypt multiple encrypted XLSX workbooks in C# with Aspose.Cells using a password list and log the outcomes
// Common Searches: C# code to decrypt a folder of password‑protected Excel files with Aspose.Cells | Aspose.Cells load encrypted XLSX workbook using multiple possible passwords in .NET | How to remove workbook password and save decrypted Excel file with Aspose.Cells | Log success and failure when batch processing protected Excel files in C# | Iterate through .xlsx files in a directory and try several passwords using Aspose.Cells
// Tags: batch decrypt xlsx Aspose.Cells | load encrypted workbook with password list | remove workbook password Aspose.Cells | save decrypted xlsx file .NET | log decryption results C#

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The example creates an output directory, scans a specified input folder for .xlsx files, and attempts to open each workbook with a list of possible passwords via Aspose.Cells LoadOptions. When the correct password is found, the workbook's password is cleared, the file is saved as a decrypted copy, and a success entry is written to a log file; failures and errors are also logged for each processed file.
class BatchDecrypt
{
    static void Main(string[] args)
    {
        // Folder containing encrypted XLSX files
        string inputFolder = @"C:\EncryptedFiles";

        // Folder where decrypted files will be saved
        string outputFolder = @"C:\DecryptedFiles";

        // Ensure output directory exists
        try
        {
            Directory.CreateDirectory(outputFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create output folder '{outputFolder}'. Exception: {ex.Message}");
            return;
        }

        // List of possible passwords to try
        List<string> passwords = new List<string>
        {
            "password1",
            "password2",
            "123456",
            "admin"
        };

        // Path for the log file
        string logPath = Path.Combine(outputFolder, "decryption_log.txt");

        try
        {
            using (StreamWriter logWriter = new StreamWriter(logPath, false))
            {
                // Verify input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    logWriter.WriteLine($"{DateTime.Now}: ERROR - Input folder '{inputFolder}' does not exist.");
                    return;
                }

                // Process each .xlsx file in the input folder
                foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    if (!File.Exists(filePath))
                    {
                        logWriter.WriteLine($"{DateTime.Now}: WARNING - File not found '{filePath}'. Skipping.");
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

                            // Remove password protection before saving
                            wb.Settings.Password = null;

                            // Save the decrypted workbook
                            string fileName = Path.GetFileNameWithoutExtension(filePath);
                            string outputPath = Path.Combine(outputFolder, fileName + "_decrypted.xlsx");
                            wb.Save(outputPath, SaveFormat.Xlsx);

                            // Log successful decryption
                            logWriter.WriteLine($"{DateTime.Now}: SUCCESS - '{filePath}' decrypted with password '{pwd}'. Saved as '{outputPath}'.");
                            decrypted = true;
                            break; // Exit password loop for this file
                        }
                        catch (CellsException ex)
                        {
                            // If the exception message indicates an incorrect password, continue trying other passwords
                            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                continue;
                            }
                            // Log other Aspose.Cells errors and stop trying passwords for this file
                            logWriter.WriteLine($"{DateTime.Now}: ERROR - '{filePath}' failed with password '{pwd}'. Exception: {ex.Message}");
                            break;
                        }
                        catch (Exception ex)
                        {
                            // Log unexpected errors and stop processing this file
                            logWriter.WriteLine($"{DateTime.Now}: ERROR - Unexpected error processing '{filePath}'. Exception: {ex.Message}");
                            break;
                        }
                    }

                    if (!decrypted)
                    {
                        // Log failure if none of the passwords worked
                        logWriter.WriteLine($"{DateTime.Now}: FAILURE - Could not decrypt '{filePath}' with provided passwords.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write log file '{logPath}'. Exception: {ex.Message}");
        }
    }
}
