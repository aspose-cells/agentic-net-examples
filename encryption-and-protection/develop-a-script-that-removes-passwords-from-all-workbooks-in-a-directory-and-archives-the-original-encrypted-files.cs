// Title: C# script to batch remove Excel workbook passwords and archive originals with Aspose.Cells
// Description: A .NET utility that scans a folder for password‑protected Excel files, copies each encrypted workbook to an "Archive" subfolder, opens the file using the known password via Aspose.Cells LoadOptions, clears the opening password, and saves the unprotected version back to its original location. Includes handling for invalid passwords and other runtime errors.
// Keywords: Aspose.Cells remove password C# | batch decrypt Excel workbooks .NET | archive encrypted Excel files | detect encrypted Excel workbook Aspose | load workbook with password Aspose.Cells | C# script Excel password removal | automate Excel file unprotect
// Common Searches: how to batch remove passwords from Excel files using Aspose.Cells | C# program to archive and decrypt protected workbooks in a directory | remove opening password from multiple Excel workbooks .NET | detect and copy encrypted Excel files before unprotecting
// Developer Intent: Programmatically strip opening passwords from every Excel workbook in a specified folder while preserving the original encrypted copies in an archive for backup or compliance purposes.
// Use Cases: Prepare incoming spreadsheets for automated reporting by removing protection in a nightly job. | Create a backup of confidential financial workbooks before bulk decryption for data analysis. | Integrate into a migration pipeline that requires unprotected Excel files for downstream processing.
// AI Prompts: Write a C# method that uses Aspose.Cells to detect encrypted Excel files in a folder, archive them, and remove their opening passwords. | Provide error‑handling code for invalid passwords when loading workbooks with Aspose.Cells LoadOptions. | Generate a PowerShell script that calls the compiled .NET executable, passes a folder path, and logs success or failure for each file.

using System;
using System.IO;
using Aspose.Cells;

// A .NET utility that scans a folder for password‑protected Excel files, copies each encrypted workbook to an "Archive" subfolder, opens the file using the known password via Aspose.Cells LoadOptions, clears the opening password, and saves the unprotected version back to its original location. Includes handling for invalid passwords and other runtime errors.
class RemoveWorkbookPasswords
{
    static void Main()
    {
        // Directory containing the workbooks
        string sourceDirectory = @"C:\Workbooks";

        // Verify source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
            return;
        }

        // Directory where original encrypted files will be archived
        string archiveDirectory = Path.Combine(sourceDirectory, "Archive");
        Directory.CreateDirectory(archiveDirectory);

        // Password used to open the encrypted workbooks (replace with actual password)
        string workbookPassword = "yourPassword";

        // Process each file in the source directory
        foreach (string filePath in Directory.GetFiles(sourceDirectory))
        {
            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Detect file format and check if the file is encrypted (password protected)
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                if (!formatInfo.IsEncrypted)
                    continue; // Skip files that are not password protected

                // Archive the original encrypted file
                string fileName = Path.GetFileName(filePath);
                string archivedPath = Path.Combine(archiveDirectory, fileName);
                File.Copy(filePath, archivedPath, overwrite: true);

                // Load the encrypted workbook using the known password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = workbookPassword
                };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Remove the workbook opening password
                workbook.Settings.Password = null;

                // Save the unprotected workbook, overwriting the original file
                workbook.Save(filePath);
                Console.WriteLine($"Processed and removed password from: {fileName}");
            }
            catch (CellsException ex)
            {
                // Handle invalid password scenario
                if (!string.IsNullOrEmpty(ex.Message) &&
                    ex.Message.IndexOf("Invalid password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Invalid password for file: {Path.GetFileName(filePath)}. Skipping.");
                }
                else
                {
                    Console.WriteLine($"CellsException processing file {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any other unexpected errors and continue processing other files
                Console.WriteLine($"Error processing file {Path.GetFileName(filePath)}: {ex.Message}");
            }
        }
    }
}
