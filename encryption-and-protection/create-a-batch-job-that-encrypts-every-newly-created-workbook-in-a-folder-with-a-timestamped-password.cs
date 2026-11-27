// Title: C# batch program to encrypt new Excel .xlsx workbooks in a folder with a timestamp‑generated password using Aspose.Cells
// Common Searches: how to encrypt all .xlsx files in a folder with a timestamp password using Aspose.Cells in C# | C# code to batch protect Excel workbooks with a datetime‑based password via Aspose.Cells | skip already password‑protected Excel files when applying encryption in a batch with Aspose.Cells | generate a timestamp string for workbook encryption in C# Aspose.Cells example | save encrypted Excel workbooks back to the original location using Aspose.Cells .NET
// Tags: batch encrypt .xlsx Aspose.Cells | timestamp password generation C# | skip existing workbook password Aspose.Cells | Workbook.Settings.Password usage | process Excel files in folder C# | encrypt newly created workbooks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified directory for .xlsx files, loads each workbook with Aspose.Cells, skips files that already require a password, creates a password from the current timestamp (yyyyMMddHHmmss), assigns it via Workbook.Settings.Password, and saves the workbook back, encrypting every newly created workbook in the folder.
class WorkbookEncryptionBatch
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

        // Get all Excel files in the folder
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx");

        foreach (string filePath in workbookFiles)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook; if it is already password‑protected, skip it
                Workbook workbook;
                try
                {
                    workbook = new Workbook(filePath);
                }
                catch (CellsException ex) when (ex.Message.Contains("Please provide password"))
                {
                    Console.WriteLine($"Skipping encrypted file (password required): {filePath}");
                    continue;
                }

                // Generate a timestamped password (e.g., 20231122153045)
                string timestampPassword = DateTime.Now.ToString("yyyyMMddHHmmss");

                // Apply the password to the workbook (encryption)
                workbook.Settings.Password = timestampPassword;

                // Save the workbook back to the same file
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
