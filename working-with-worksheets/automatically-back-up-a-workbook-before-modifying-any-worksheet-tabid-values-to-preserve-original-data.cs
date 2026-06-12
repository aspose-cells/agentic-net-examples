using System;
using System.IO;
using Aspose.Cells;

class BackupAndModifyTabId
{
    static void Main()
    {
        // Paths for the original, backup, and modified files
        string sourcePath = "input.xlsx";
        string backupPath = "input_backup.xlsx";
        string outputPath = "input_modified.xlsx";

        // Verify that the source file exists
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the original workbook
            Workbook workbook = new Workbook(sourcePath);

            // Create a backup copy of the workbook
            Workbook backupWorkbook = new Workbook();
            backupWorkbook.Copy(workbook);               // Deep copy of the workbook
            backupWorkbook.Save(backupPath);              // Save the backup

            // Modify the TabId of each worksheet (example: assign a new random identifier)
            Random rnd = new Random();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.TabId = rnd.Next(1, 1000);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine("Backup and modification completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}