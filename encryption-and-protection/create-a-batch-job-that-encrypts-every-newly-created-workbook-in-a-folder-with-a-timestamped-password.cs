// Title: C# Batch Encrypt Excel Files with Timestamp Passwords via Aspose.Cells
// Description: Scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, generates a yyyyMMddHHmmss timestamp password, applies it via Workbook.Settings.Password, saves the file, logs the password and skips already‑protected workbooks.
// Keywords: Aspose.Cells C# encryption | batch encrypt Excel files | timestamp password Excel | protect multiple workbooks programmatically | skip already protected Excel | folder based workbook security | automated Excel file encryption | C# Aspose.Cells example | GitHub Aspose.Cells batch encryption | Excel password protection script
// Common Searches: how to encrypt all Excel files in a folder using Aspose.Cells C# | timestamp based password for Excel workbook Aspose.Cells | batch protect newly created workbooks C# | skip already password protected Excel files Aspose.Cells | automate Excel file encryption with timestamp
// Developer Intent: Encrypt every workbook placed in a specific directory by assigning a unique timestamp‑derived password with Aspose.Cells.
// Use Cases: Secure daily generated reports before archiving on shared storage. | Automate protection of exported spreadsheets in a data‑processing pipeline. | Create an audit trail by logging the timestamp password for each encrypted file.
// AI Prompts: Write C# code that watches a folder and encrypts any new .xlsx file with a timestamp password using Aspose.Cells. | Provide an Aspose.Cells example that batch encrypts all Excel files in a directory and writes each file's password to a CSV log. | Show how to skip already password‑protected workbooks while encrypting a folder of Excel files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchEncryption
{
    // Scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, generates a yyyyMMddHHmmss timestamp password, applies it via Workbook.Settings.Password, saves the file, logs the password and skips already‑protected workbooks.
    class Program
    {
        static void Main()
        {
            // Folder containing the workbooks to encrypt
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files in the folder (you can add other extensions if needed)
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the existing workbook
                    Workbook workbook = new Workbook(filePath);

                    // Generate a timestamped password (e.g., 20230815103045)
                    string timestampPassword = DateTime.Now.ToString("yyyyMMddHHmmss");

                    // Set the password for the workbook (encryption)
                    workbook.Settings.Password = timestampPassword;

                    // Overwrite the original file with the encrypted version
                    workbook.Save(filePath);

                    // Optional: output the applied password for logging purposes
                    Console.WriteLine($"Encrypted '{Path.GetFileName(filePath)}' with password: {timestampPassword}");
                }
                catch (CellsException ex)
                {
                    // If the workbook is already password‑protected, Aspose.Cells throws a CellsException.
                    // We treat this as a skip scenario.
                    Console.WriteLine($"Skipped already protected file: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Log any other unexpected errors and continue processing remaining files
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch encryption completed.");
        }
    }
}
