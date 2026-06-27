using System;
using Aspose.Cells;

namespace BackupAndModifyExternalLinks
{
    class Program
    {
        static void Main()
        {
            // Path to the original workbook
            string originalPath = "OriginalWorkbook.xlsx";

            // Load the original workbook (load rule)
            Workbook originalWorkbook = new Workbook(originalPath);

            // Create a backup workbook (create rule)
            Workbook backupWorkbook = new Workbook();

            // Copy the content of the original workbook to the backup (copy rule)
            backupWorkbook.Copy(originalWorkbook);

            // Save the backup workbook (save rule)
            string backupPath = "OriginalWorkbook_Backup.xlsx";
            backupWorkbook.Save(backupPath);

            // Modify external link paths in the original workbook
            ExternalLinkCollection externalLinks = originalWorkbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                // Example modification: replace a folder name in the data source
                string currentSource = externalLinks[i].DataSource;
                string modifiedSource = currentSource.Replace(@"C:\OldFolder\", @"D:\NewFolder\");
                externalLinks[i].DataSource = modifiedSource;
            }

            // Save the modified original workbook (save rule)
            string modifiedPath = "OriginalWorkbook_Modified.xlsx";
            originalWorkbook.Save(modifiedPath);

            // Clean up
            originalWorkbook.Dispose();
            backupWorkbook.Dispose();

            Console.WriteLine("Backup created at: " + backupPath);
            Console.WriteLine("Modified workbook saved at: " + modifiedPath);
        }
    }
}