// Title: Batch remove Excel passwords and archive originals using Aspose.Cells C#
// Description: A C# console program that scans a folder, detects encrypted Excel workbooks with Aspose.Cells, copies each protected file to an "Archive" subfolder, opens the workbook using a known password, clears the password, and saves the file back in its original format while preserving a backup of the encrypted version.
// Keywords: Aspose.Cells password removal C# | batch unprotect Excel files .NET | archive encrypted workbooks | detect encrypted Excel workbook | remove workbook password Aspose | C# script Excel security automation
// Common Searches: how to batch remove passwords from Excel files using Aspose.Cells | C# script to archive and unprotect encrypted workbooks | remove Excel workbook password programmatically .NET | detect and decrypt password‑protected Excel files C#
// Developer Intent: Automatically strip passwords from every encrypted workbook in a directory while keeping a backup of the original protected files.
// Use Cases: Prepare a shared reports folder for downstream analytics by unprotecting spreadsheets but retain the original encrypted copies for audit compliance. | Run a nightly job that removes passwords from incoming Excel files so downstream tools can process them, with the originals archived for legal retention. | Create a secure archive of all password‑protected workbooks before converting them to an unprotected state for bulk editing or migration.
// AI Prompts: Generate a C# console app that uses Aspose.Cells to detect encrypted Excel files in a folder, copy them to an Archive subfolder, open each with a known password, remove the password, and save the unprotected workbook. | Show how to extend the script to support a dictionary of filenames and their respective passwords for batch decryption. | Provide sample logging code that records each file’s encryption status, archive path, and success/failure of password removal.

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that scans a folder, detects encrypted Excel workbooks with Aspose.Cells, copies each protected file to an "Archive" subfolder, opens the workbook using a known password, clears the password, and saves the file back in its original format while preserving a backup of the encrypted version.
class RemoveWorkbookPasswords
{
    static void Main()
    {
        // Directory containing the workbooks
        string sourceDirectory = @"C:\Workbooks";

        // Directory where original encrypted files will be archived
        string archiveDirectory = Path.Combine(sourceDirectory, "Archive");
        Directory.CreateDirectory(archiveDirectory);

        // Password used to open the encrypted workbooks (replace with actual password)
        const string workbookPassword = "password";

        // Process each file in the source directory
        foreach (string filePath in Directory.GetFiles(sourceDirectory))
        {
            // Detect file format and check if the file is encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            if (!formatInfo.IsEncrypted)
                continue; // Skip files that are not password protected

            // Archive the original encrypted file before modifying it
            string archivePath = Path.Combine(archiveDirectory, Path.GetFileName(filePath));
            File.Copy(filePath, archivePath, true);

            // Load the encrypted workbook using the known password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.Password = workbookPassword;
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Remove the password protection
            workbook.Settings.Password = null;

            // Save the workbook back to the original location (unprotected)
            // Save without specifying format to keep the original file type
            workbook.Save(filePath);
        }

        Console.WriteLine("Password removal completed. Original encrypted files are archived.");
    }
}
