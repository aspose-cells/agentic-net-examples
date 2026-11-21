// Title: Backup worksheets with a specific name prefix to a new Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells copy only worksheets that start with 'Report_' to another workbook in C# | How to create a backup Excel file containing selected sheets by name prefix using Aspose.Cells .NET | C# code to filter worksheets by prefix and save them into a new workbook with Aspose.Cells | Save subset of worksheets to a new .xlsx file using Aspose.Cells AddCopy method | Backup specific sheets from source.xlsx to backup.xlsx based on sheet name prefix in C#
// Tags: Aspose.Cells AddCopy worksheet copying | C# create backup workbook from selected sheets | select worksheets by name prefix Aspose.Cells | save subset of sheets to new .xlsx using Aspose.Cells | ensure output directory exists before saving workbook C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source Excel file, creates an empty workbook, iterates through all worksheets, and uses Aspose.Cells' AddCopy method to copy only those whose names start with a given prefix. It ensures the target directory exists, saves the new workbook as a backup, and reports success or errors.
class BackupWorksheets
{
    static void Main()
    {
        try
        {
            // Path to the original workbook
            string sourcePath = "source.xlsx";

            // Prefix to filter worksheets (e.g., "Report_")
            string prefix = "Report_";

            // Path for the backup workbook
            string backupPath = "backup.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the original workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook for backup
            Workbook backupWorkbook = new Workbook();

            // Remove the default empty worksheet created with the new workbook
            backupWorkbook.Worksheets.Clear();

            // Iterate through all worksheets in the source workbook
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Check if the worksheet name starts with the specified prefix
                if (sheet.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    // Copy the matching worksheet into the backup workbook
                    // AddCopy expects the worksheet name, not the Worksheet object
                    backupWorkbook.Worksheets.AddCopy(sheet.Name);
                }
            }

            // Ensure the directory for the backup file exists
            string backupDir = Path.GetDirectoryName(backupPath);
            if (!string.IsNullOrEmpty(backupDir) && !Directory.Exists(backupDir))
                Directory.CreateDirectory(backupDir);

            // Save the backup workbook to the specified file
            backupWorkbook.Save(backupPath);
            Console.WriteLine($"Backup created successfully at '{backupPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
