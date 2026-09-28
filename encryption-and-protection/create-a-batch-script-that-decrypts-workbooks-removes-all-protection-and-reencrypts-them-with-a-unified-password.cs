// Title: C# batch script to decrypt Excel workbooks, strip all protection, and re‑encrypt them with a single password using Aspose.Cells
// AI Prompts: Generate a C# console application that recursively scans a source folder, loads each Excel file with an optional existing password, calls Workbook.Unprotect and Worksheet.Unprotect, assigns a new password via Workbook.Settings.Password, and saves the file to a target folder while preserving the original directory hierarchy. | Enhance the program to catch password‑mismatch exceptions, skip those files, and log the file paths that could not be processed. | Add a final summary that reports the total number of files examined, how many were successfully re‑encrypted, and how many failed.
// Common Searches: aspocells batch remove protection from multiple Excel workbooks c# | change password of encrypted .xlsx files programmatically using Aspose.Cells | re‑encrypt a folder of Excel files with a new password .NET console app | preserve folder structure when saving processed Excel workbooks c# | skip files with wrong original password Aspose.Cells LoadOptions
// Tags: decrypt and re‑encrypt Excel workbooks Aspose.Cells | remove workbook and worksheet protection .NET | batch process Excel files recursively | set unified password for encrypted workbooks | load encrypted workbook with LoadOptions

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookBatchProcessor
{
    // The program iterates through all Excel files (.xlsx, .xls, .xlsm, .xlsb) in a source directory, optionally opens them with an existing password, removes workbook and worksheet protection, applies a new unified password, and saves the files to a target directory while maintaining the original folder hierarchy, using Aspose.Cells for .NET.
    class Program
    {
        // Entry point
        static void Main(string[] args)
        {
            // args[0] - source folder
            // args[1] - target folder
            // args[2] - original password (empty if none)
            // args[3] - new password to apply
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: WorkbookBatchProcessor <sourceFolder> <targetFolder> <originalPassword> <newPassword>");
                return;
            }

            string sourceFolder = args[0];
            string targetFolder = args[1];
            string originalPassword = args[2];
            string newPassword = args[3];

            // Ensure target folder exists
            Directory.CreateDirectory(targetFolder);

            // Get all files (including subfolders)
            string[] files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories);
            foreach (string filePath in files)
            {
                // Process only supported Excel formats
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                    continue;

                // Verify the file actually exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // ---------- Load ----------
                    var loadOptions = new LoadOptions(LoadFormat.Auto);
                    if (!string.IsNullOrEmpty(originalPassword))
                        loadOptions.Password = originalPassword;

                    var workbook = new Workbook(filePath, loadOptions);

                    // ---------- Remove Protection ----------
                    // Unprotect workbook (no need to check IsProtected)
                    workbook.Unprotect(originalPassword);

                    // Unprotect each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.Unprotect(originalPassword);
                    }

                    // ---------- Apply New Encryption ----------
                    workbook.Settings.Password = newPassword;

                    // ---------- Save ----------
                    string relativePath = Path.GetRelativePath(sourceFolder, filePath);
                    string outputPath = Path.Combine(targetFolder, relativePath);
                    string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                    Directory.CreateDirectory(outputDir);

                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Processed: {relativePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
