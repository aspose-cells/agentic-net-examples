// Title: Batch remove opening passwords from Excel .xlsx workbooks and archive the original encrypted files with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a folder for *.xlsx files, copies each file to an Archive subfolder, opens the workbook with a known password using Aspose.Cells LoadOptions, clears the workbook password, and saves the file back without protection. | Modify the script to accept the source directory and password as command‑line arguments and log any files that fail to open due to an incorrect password into a separate log file. | Extend the solution to process sub‑directories recursively and generate a summary report showing how many workbooks were successfully decrypted versus how many failed.
// Common Searches: how to remove password protection from multiple Excel files using Aspose.Cells in C# | C# script to copy encrypted .xlsx files to archive folder before clearing workbook password | batch processing of Excel workbooks to strip opening password with Aspose.Cells LoadOptions | automate removal of workbook password and preserve original files in .NET | Aspose.Cells remove workbook opening password for all files in a directory
// Tags: batch remove Excel workbook passwords Aspose.Cells | archive original encrypted .xlsx files C# | encrypted workbook loading Aspose.Cells | save workbook without opening protection .NET | recursive Excel password removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# console app iterates over all .xlsx files in a specified folder, copies each encrypted workbook to an Archive subdirectory, opens each file with a known password via Aspose.Cells LoadOptions, clears the opening password, saves the workbook unprotected, and logs any files that cannot be processed.
class RemoveWorkbookPasswords
{
    static void Main()
    {
        // Directory containing the encrypted workbooks
        string sourceDirectory = @"C:\Workbooks";

        // Ensure the source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
            return;
        }

        // Directory where original encrypted files will be archived
        string archiveDirectory = Path.Combine(sourceDirectory, "Archive");
        Directory.CreateDirectory(archiveDirectory);

        // Known password for the encrypted workbooks (adjust as needed)
        const string knownPassword = "password";

        // Process each Excel file in the source directory
        foreach (string filePath in Directory.GetFiles(sourceDirectory, "*.xlsx"))
        {
            // Verify the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Preserve the original encrypted file by copying it to the archive folder
                string fileName = Path.GetFileName(filePath);
                string archivePath = Path.Combine(archiveDirectory, fileName);
                File.Copy(filePath, archivePath, true);

                // Load the workbook using the known password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = knownPassword
                };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Remove the opening password
                workbook.Settings.Password = null;

                // Save the workbook without a password, overwriting the original file
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                // Log or handle files that could not be processed (e.g., wrong password)
                Console.WriteLine($"Failed to process '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Password removal complete. Original files archived.");
    }
}
