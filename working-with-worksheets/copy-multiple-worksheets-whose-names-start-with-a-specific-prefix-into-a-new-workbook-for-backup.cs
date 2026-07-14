using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorksheetBackupDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            string sourcePath = "source.xlsx";
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define the prefix to filter worksheets
            string prefix = "Report_";

            // Collect worksheets whose names start with the specified prefix
            List<Worksheet> sheetsToCopy = new List<Worksheet>();
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                if (ws.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    sheetsToCopy.Add(ws);
                }
            }

            // Create a new workbook for the backup
            Workbook backupWorkbook = new Workbook();

            // Remove the default sheet that comes with a new workbook
            backupWorkbook.Worksheets.Clear();

            // Copy each selected worksheet into the backup workbook
            foreach (Worksheet sourceSheet in sheetsToCopy)
            {
                // Add a new blank worksheet with the same name as the source
                Worksheet destSheet = backupWorkbook.Worksheets.Add(sourceSheet.Name);

                // Copy contents and formats from the source worksheet
                destSheet.Copy(sourceSheet);
            }

            // Save the backup workbook
            string backupPath = "backup.xlsx";
            backupWorkbook.Save(backupPath);

            Console.WriteLine($"Backup completed. {backupWorkbook.Worksheets.Count} worksheets saved to '{backupPath}'.");
        }
    }
}