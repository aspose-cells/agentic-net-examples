using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace BackupWorksheetsExample
{
    class Program
    {
        static void Main()
        {
            // Define the prefix to filter worksheets
            const string prefix = "Report_";

            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Collect worksheets whose names start with the specified prefix
            List<Worksheet> matchingSheets = new List<Worksheet>();
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                if (sheet.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    matchingSheets.Add(sheet);
                }
            }

            // If no matching worksheets are found, exit
            if (matchingSheets.Count == 0)
            {
                Console.WriteLine("No worksheets found with the specified prefix.");
                return;
            }

            // Prepare source array and destination names array
            Worksheet[] sourceArray = matchingSheets.ToArray();
            string[] destNames = sourceArray.Select(ws => ws.Name).ToArray();

            // Create a new workbook for backup
            Workbook backupWorkbook = new Workbook();

            // Remove the default sheet that Aspose.Cells creates automatically
            backupWorkbook.Worksheets.Clear();

            // Copy the selected worksheets into the backup workbook
            backupWorkbook.Worksheets.AddCopy(sourceArray, destNames);

            // Save the backup workbook
            backupWorkbook.Save("backup.xlsx");

            Console.WriteLine("Backup completed successfully.");
        }
    }
}