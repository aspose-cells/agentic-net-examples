// Title: Backup an Excel workbook and adjust worksheet TabId values with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, creates a copy as a backup file, increments the TabId of every worksheet, and saves both the backup and the modified original workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# backup workbook | Excel workbook copy | Worksheet TabId | modify TabId Aspose.Cells | preserve original Excel file | copy workbook C# | save backup Excel | Aspose.Cells .NET example
// Common Searches: Aspose.Cells backup workbook before editing | How to copy an Excel file with Aspose.Cells C# | Change worksheet TabId using Aspose.Cells | Save original Excel file as backup in .NET | C# example for preserving workbook data with Aspose
// Developer Intent: Create a backup of the original Excel file, then modify each worksheet's TabId while keeping the source data unchanged.
// Use Cases: Automated data‑migration scripts that need a safety copy before altering sheet properties. | Audit‑trail generation where a snapshot of the workbook is stored prior to batch updates. | Scheduled maintenance jobs that back up workbooks and then re‑index TabId values for integration with external systems.
// AI Prompts: Generate C# code using Aspose.Cells to copy an existing workbook to a backup file, increase every worksheet's TabId by a specified offset, and save both files. | Provide a reusable method that accepts a workbook path and an offset, creates a backup with a timestamped suffix, updates TabId values for all worksheets, and returns the paths of the backup and updated files.

using System;
using Aspose.Cells;

namespace WorkbookBackupExample
{
    // Loads an existing workbook, creates a copy as a backup file, increments the TabId of every worksheet, and saves both the backup and the modified original workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the original workbook
            string originalPath = "OriginalWorkbook.xlsx";

            // Load the original workbook (load rule)
            Workbook originalWorkbook = new Workbook(originalPath);

            // Create a backup workbook instance (create rule)
            Workbook backupWorkbook = new Workbook();

            // Copy the original workbook into the backup workbook (copy rule)
            backupWorkbook.Copy(originalWorkbook);

            // Save the backup workbook (save rule)
            string backupPath = "OriginalWorkbook_Backup.xlsx";
            backupWorkbook.Save(backupPath);

            // Modify TabId values of each worksheet in the original workbook
            foreach (Worksheet sheet in originalWorkbook.Worksheets)
            {
                // Example modification: set TabId to a new unique value
                // Here we simply add 1000 to the existing TabId
                sheet.TabId += 1000;
            }

            // Save the modified original workbook (save rule)
            originalWorkbook.Save(originalPath);
        }
    }
}
