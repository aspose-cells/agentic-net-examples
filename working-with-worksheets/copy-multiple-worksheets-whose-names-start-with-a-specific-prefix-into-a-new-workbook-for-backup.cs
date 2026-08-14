// Title: C# – Copy Worksheets with a Prefix to a New Workbook (Backup) using Aspose.Cells
// Description: Loads a source Excel file, selects worksheets whose names start with a specified prefix (case‑insensitive), copies them as a group into an empty workbook via AddCopy, and saves the result as a backup file.
// Keywords: Aspose.Cells C# copy worksheets | Excel sheet prefix filter | backup selected worksheets | AddCopy multiple sheets | export worksheets by name | filter worksheets Aspose.Cells | create backup workbook .NET | copy sheets with specific prefix | Aspose.Cells worksheet selection | C# Excel automation backup
// Common Searches: Aspose.Cells copy worksheets by prefix | C# backup selected Excel sheets using Aspose | How to copy multiple sheets to a new workbook Aspose.Cells | filter worksheets that start with Report_ Aspose.Cells | AddCopy example for copying several worksheets
// Developer Intent: Select all worksheets whose names begin with a given prefix and copy them into a new workbook for backup purposes.
// Use Cases: Create a nightly backup that includes only report sheets named with a common prefix. | Export a subset of worksheets matching a naming convention to a separate file for client delivery. | Archive version‑specific sheets before applying bulk changes to the original workbook.
// AI Prompts: Write C# code with Aspose.Cells that copies every worksheet whose name starts with 'Report_' into a new workbook called backup.xlsx. | Show how to filter worksheets by a case‑insensitive prefix and copy them as a group while preserving original sheet names using Aspose.Cells. | Explain how to modify the example to add a '_Backup' suffix to each copied worksheet name without altering the source workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorksheetBackupExample
{
    // Loads a source Excel file, selects worksheets whose names start with a specified prefix (case‑insensitive), copies them as a group into an empty workbook via AddCopy, and saves the result as a backup file.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook for backup
            Workbook backupWorkbook = new Workbook();

            // Remove the default worksheet that Aspose.Cells creates
            backupWorkbook.Worksheets.Clear();

            // Define the prefix to filter worksheets
            string prefix = "Report_";

            // Collect source worksheets and corresponding destination names
            List<Worksheet> sourceSheets = new List<Worksheet>();
            List<string> destNames = new List<string>();

            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                if (sheet.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    sourceSheets.Add(sheet);
                    // Use the same name in the backup workbook (or modify as needed)
                    destNames.Add(sheet.Name);
                }
            }

            // If there are matching worksheets, copy them as a group
            if (sourceSheets.Count > 0)
            {
                // AddCopy(Worksheet[], string[]) copies the group of worksheets
                backupWorkbook.Worksheets.AddCopy(sourceSheets.ToArray(), destNames.ToArray());
            }

            // Save the backup workbook
            backupWorkbook.Save("backup.xlsx");
        }
    }
}
