// Title: Backup an Excel workbook and modify worksheet TabId using Aspose.Cells for .NET
// Description: Load a workbook, create a backup with Workbook.Copy, save it, update each worksheet's TabId, and write the modified file to a new location—all with Aspose.Cells in C#.
// Keywords: Aspose.Cells backup workbook | C# modify worksheet TabId | Workbook.Copy example | save Excel file after changes .NET | preserve original Excel data | Aspose.Cells Worksheet.TabId | Excel versioning C#
// Common Searches: Aspose.Cells create backup before editing workbook | How to change TabId of worksheets in C# | Copy Excel file with Aspose.Cells and keep original | C# backup Excel workbook then modify sheets | Aspose.Cells example for preserving original data
// Developer Intent: Generate a backup of the source workbook, change each sheet's TabId, and save the altered workbook without overwriting the original.
// Use Cases: Automated service that backs up user‑uploaded spreadsheets before applying bulk TabId updates. | Version‑controlled Excel processing pipeline where the original file must remain unchanged. | Data‑integrity safeguard for scripts that reorder or rename worksheets by adjusting TabId values.
// AI Prompts: Write C# code with Aspose.Cells to copy a workbook to a backup file, update each worksheet's TabId, and save the result to a new file. | Show how to add error handling and logging around the backup and TabId modification steps using Aspose.Cells. | Explain how to compare the backup workbook with the source to confirm they are identical before modifying TabId values.

using System;
using Aspose.Cells;

namespace AsposeCellsBackupDemo
{
    // Load a workbook, create a backup with Workbook.Copy, save it, update each worksheet's TabId, and write the modified file to a new location—all with Aspose.Cells in C#.
    public class TabIdModifier
    {
        /// <param name="sourcePath">Path to the original workbook.</param>
        /// <param name="backupPath">Path where the backup will be saved.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        public static void Process(string sourcePath, string backupPath, string outputPath)
        {
            // Load the original workbook (uses the provided Workbook(string) constructor)
            Workbook originalWorkbook = new Workbook(sourcePath);

            // Create a backup workbook by copying the original (uses Workbook.Copy method)
            Workbook backupWorkbook = new Workbook();
            backupWorkbook.Copy(originalWorkbook);

            // Save the backup (uses the provided Save(string) method)
            backupWorkbook.Save(backupPath);

            // Modify TabId values of each worksheet in the original workbook
            foreach (Worksheet sheet in originalWorkbook.Worksheets)
            {
                // Example modification: set TabId to a unique value based on sheet index
                sheet.TabId = sheet.Index + 1000;
            }

            // Save the modified workbook (uses the provided Save(string) method)
            originalWorkbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "InputWorkbook.xlsx";
            string backupFile = "InputWorkbook_Backup.xlsx";
            string outputFile = "InputWorkbook_Modified.xlsx";

            Process(sourceFile, backupFile, outputFile);

            Console.WriteLine("Backup created at: " + backupFile);
            Console.WriteLine("Modified workbook saved at: " + outputFile);
        }
    }
}
