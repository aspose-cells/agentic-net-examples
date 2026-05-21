using System;
using Aspose.Cells;

class BackupAndModifyExternalLinks
{
    static void Main()
    {
        // Load the original workbook
        string sourcePath = "input.xlsx";
        Workbook originalWorkbook = new Workbook(sourcePath);

        // Create a backup copy of the workbook
        Workbook backupWorkbook = new Workbook();
        backupWorkbook.Copy(originalWorkbook);               // use the Copy method (create‑>copy)
        backupWorkbook.Save("backup_input.xlsx");            // save the backup

        // Modify external link paths in the original workbook
        ExternalLinkCollection externalLinks = originalWorkbook.Worksheets.ExternalLinks;
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Example modification: replace an old folder path with a new one
            string oldPath = link.DataSource;
            string newPath = oldPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");
            link.DataSource = newPath;
        }

        // Save the workbook after modifications
        originalWorkbook.Save("output_modified.xlsx");
    }
}