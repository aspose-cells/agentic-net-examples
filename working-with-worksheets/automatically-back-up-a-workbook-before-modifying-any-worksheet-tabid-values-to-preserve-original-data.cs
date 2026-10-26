// Title: Create a backup of an Excel workbook before updating worksheet TabId values with Aspose.Cells in C#
// Common Searches: c# asp.net create backup of Excel workbook before changing worksheet tab identifiers with Aspose.Cells | how to preserve original Excel file while updating worksheet TabId using Aspose.Cells library | example code to save a copy of a workbook and then set TabId values in C# | aspocells backup workbook then modify worksheet properties programmatically | C# Aspose.Cells modify worksheet TabId after creating backup copy
// Tags: Aspose.Cells generate backup copy of workbook | set worksheet TabId using Aspose.Cells C# | Workbook.Save duplicate file programmatically | modify worksheet properties with Aspose.Cells | retain source workbook during Excel modifications

using Aspose.Cells;
using System;

// The sample loads an existing Excel file, creates a duplicate backup using Workbook.Save, updates each worksheet's TabId property, and then saves the altered workbook as a new file, ensuring the original workbook remains unchanged.
class Program
{
    static void Main()
    {
        // Path to the original workbook
        string originalPath = "original.xlsx";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(originalPath);

        // Create a backup copy before making any changes (save rule)
        string backupPath = "original_backup.xlsx";
        workbook.Save(backupPath);

        // Modify TabId values for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Example: set TabId to a new unique value (e.g., current index + 100)
            sheet.TabId = sheet.Index + 100;
        }

        // Save the modified workbook (save rule)
        string modifiedPath = "original_modified.xlsx";
        workbook.Save(modifiedPath);
    }
}
