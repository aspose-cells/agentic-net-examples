// Title: How to create a timestamped backup of an Excel workbook before updating worksheet TabId values with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# backup original Excel file before changing worksheet TabId | Create a copy of a workbook with a timestamp using Aspose.Cells .NET | Preserve original workbook when modifying worksheet TabId in C# | C# program to backup Excel file then set new TabId for each sheet with Aspose.Cells | How to copy an Excel workbook to a backup folder before editing with Aspose.Cells
// Tags: automated workbook backup Aspose.Cells | set worksheet TabId C# Aspose.Cells | preserve original Excel file Aspose.Cells | save modified workbook to new file Aspose.Cells | create backup directory .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookTabIdBackupDemo
{
    // The sample validates input arguments, determines output and backup locations, creates a timestamped backup copy of the original workbook, loads the workbook with Aspose.Cells, iterates through each worksheet to assign new TabId values, and saves the modified workbook to a separate file while handling exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args.Length < 1)
                {
                    Console.WriteLine("Error: Input workbook path is required.");
                    Console.WriteLine("Usage: WorkbookTabIdBackupDemo.exe <inputWorkbookPath> [outputWorkbookPath] [backupDirectory]");
                    return;
                }

                string inputPath = args[0];
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                // Determine output path
                string outputPath;
                if (args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1]))
                {
                    outputPath = args[1];
                }
                else
                {
                    string dir = Path.GetDirectoryName(inputPath);
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                    string ext = Path.GetExtension(inputPath);
                    outputPath = Path.Combine(dir, $"{nameWithoutExt}_modified{ext}");
                }

                // Determine backup directory
                string backupDir;
                if (args.Length >= 3 && !string.IsNullOrWhiteSpace(args[2]))
                {
                    backupDir = args[2];
                }
                else
                {
                    backupDir = Path.GetDirectoryName(inputPath);
                }

                // Ensure backup directory exists
                if (!Directory.Exists(backupDir))
                {
                    Directory.CreateDirectory(backupDir);
                }

                // Create backup file name with timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupFileName = $"{Path.GetFileNameWithoutExtension(inputPath)}_backup_{timestamp}{Path.GetExtension(inputPath)}";
                string backupPath = Path.Combine(backupDir, backupFileName);

                // Perform the backup (copy the original file)
                File.Copy(inputPath, backupPath, overwrite: true);
                Console.WriteLine($"Backup created at: {backupPath}");

                // Load the workbook from the original file (not the backup)
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook loaded. Worksheet count: {workbook.Worksheets.Count}");

                // Modify TabId for each worksheet
                // Example strategy: set TabId = worksheet index + 100
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    int originalTabId = ws.TabId;
                    ws.TabId = i + 100; // arbitrary new value
                    Console.WriteLine($"Worksheet \"{ws.Name}\": TabId changed from {originalTabId} to {ws.TabId}");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Modified workbook saved at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
