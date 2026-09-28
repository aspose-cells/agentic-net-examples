// Title: Create a backup of an Excel workbook and change all worksheet TabId values using Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies an existing .xlsx file to a backup location, then iterates through each worksheet and sets its TabId to Index+1 using Aspose.Cells. | Show how to create a minimal workbook when the source file is missing, ensure the backup and output directories exist, and save the modified workbook to a separate path with Aspose.Cells.
// Common Searches: aspocells c# backup workbook before changing worksheet TabId | set worksheet TabId programmatically with Aspose.Cells after creating backup | C# example to copy Excel file, modify sheet TabId, and save to new location using Aspose.Cells
// Tags: Aspose.Cells workbook backup creation | Aspose.Cells modify worksheet TabId | Aspose.Cells save workbook to new file path | Aspose.Cells ensure output directory exists | Aspose.Cells generate placeholder workbook if missing

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to verify an Excel file's existence, create a minimal workbook if needed, back up the original workbook, update each worksheet's TabId to its index plus one, and save the modified workbook to a separate location using Aspose.Cells for .NET.
class WorkbookBackupAndModify
{
    static void Main()
    {
        // Paths for the original workbook, backup copy, and the modified workbook
        string originalPath = @"C:\Data\OriginalWorkbook.xlsx";
        string backupPath   = @"C:\Data\Backup\OriginalWorkbook_Backup.xlsx";
        string outputPath   = @"C:\Data\Modified\OriginalWorkbook_Modified.xlsx";

        try
        {
            // Ensure the original file exists; if not, create a minimal workbook
            if (!File.Exists(originalPath))
            {
                // Create directory if needed
                Directory.CreateDirectory(Path.GetDirectoryName(originalPath));

                // Create a new workbook with a single worksheet and save it as the original file
                var newWb = new Workbook();
                newWb.Worksheets[0].Name = "Sheet1";
                newWb.Save(originalPath, SaveFormat.Xlsx);
            }

            // Load the original workbook
            Workbook workbook = new Workbook(originalPath);

            // Ensure backup and output directories exist
            Directory.CreateDirectory(Path.GetDirectoryName(backupPath));
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Create a backup of the workbook before any changes
            workbook.Save(backupPath, SaveFormat.Xlsx);

            // Modify the TabId of each worksheet as needed (example: index + 1)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // TabId must be a positive integer; adjust as required
                sheet.TabId = sheet.Index + 1;
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
