// Title: Encrypt newly created Excel workbooks in a folder with timestamp‑based passwords using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a given directory for .xlsx files, loads each workbook with Aspose.Cells, assigns a password generated from the current timestamp, and saves the workbook encrypted. | Update the batch encryption script to detect workbooks that already have a password and skip them, while still applying a timestamp‑derived password to unprotected files. | Add logging to the program so that each encrypted workbook’s filename and generated password are written to a log file, and ensure exceptions are handled gracefully.
// Common Searches: how to use Aspose.Cells to add a timestamp password to multiple Excel files in C# | C# program to encrypt all .xlsx files in a folder with a dynamic password | batch encrypt Excel workbooks with Aspose.Cells and save them back to the same location | automate Excel file protection using current date and time as password in .NET
// Tags: Aspose.Cells workbook password protection | dynamic date‑based password C# | batch encrypt .xlsx files .NET | programmatic Excel file security Aspose | folder scan encrypt Excel workbooks

using System;
using System.IO;
using Aspose.Cells;

// // This console app iterates over every .xlsx file in a specified directory, loads each workbook with Aspose.Cells, sets a password derived from the current timestamp, and saves the file back encrypted.
class WorkbookEncryptionBatch
{
    static void Main()
    {
        // Folder containing the workbooks to encrypt
        string folderPath = @"C:\Workbooks";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all Excel files in the folder
        string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Generate a timestamped password (e.g., 20230920143055)
                string timestampPassword = DateTime.Now.ToString("yyyyMMddHHmmss");

                // Apply password protection to the workbook (file encryption)
                workbook.Settings.Password = timestampPassword;

                // Save the workbook back to the same file, now encrypted
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Encrypted: {Path.GetFileName(filePath)} with password {timestampPassword}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
