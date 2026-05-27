using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

class BatchDecryptSanitizeEncrypt
{
    // Entry point
    static void Main(string[] args)
    {
        // Example usage:
        // args[0] = input folder path
        // args[1] = output folder path
        // args[2] = old password (used for all encrypted files)
        // args[3] = new password (to apply after sanitization)

        if (args.Length < 4)
        {
            Console.WriteLine("Usage: BatchDecryptSanitizeEncrypt <inputFolder> <outputFolder> <oldPassword> <newPassword>");
            return;
        }

        string inputFolder = args[0];
        string outputFolder = args[1];
        string oldPassword = args[2];
        string newPassword = args[3];

        ProcessFolder(inputFolder, outputFolder, oldPassword, newPassword);
    }

    // Process all Excel files in a folder (non‑recursive)
    static void ProcessFolder(string inputFolder, string outputFolder, string oldPwd, string newPwd)
    {
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Supported Excel extensions
        string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods" };

        foreach (string filePath in Directory.GetFiles(inputFolder))
        {
            if (Array.Exists(extensions, e => e.Equals(Path.GetExtension(filePath), StringComparison.OrdinalIgnoreCase)))
            {
                try
                {
                    string fileName = Path.GetFileName(filePath);
                    string outPath = Path.Combine(outputFolder, fileName);
                    ProcessFile(filePath, outPath, oldPwd, newPwd);
                    Console.WriteLine($"Processed: {fileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {filePath}: {ex.Message}");
                }
            }
        }
    }

    // Decrypt, sanitize, and re‑encrypt a single workbook
    static void ProcessFile(string sourcePath, string destPath, string oldPwd, string newPwd)
    {
        // Detect file format to know if it is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);

        Workbook workbook;

        if (formatInfo.IsEncrypted)
        {
            // Load encrypted workbook using the old password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = oldPwd;
            workbook = new Workbook(sourcePath, loadOptions);
        }
        else
        {
            // Load normally
            workbook = new Workbook(sourcePath);
        }

        // ---------- Data sanitization ----------
        // Remove personal information (comments author, document properties, etc.)
        workbook.RemovePersonalInformation();

        // Optionally clear all comments (example of deeper sanitization)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Comments.Clear();
        }

        // ---------- Re‑encryption ----------
        // Set the new password for the workbook
        workbook.Settings.Password = newPwd;

        // Save the sanitized and re‑encrypted workbook
        workbook.Save(destPath);
    }
}