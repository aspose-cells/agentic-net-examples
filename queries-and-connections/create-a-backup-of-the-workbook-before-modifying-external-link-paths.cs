using System;
using Aspose.Cells;

class BackupAndModifyExternalLinks
{
    static void Main()
    {
        // Load the original workbook
        string originalFile = "input.xlsx";
        Workbook workbook = new Workbook(originalFile);

        // Save a backup copy before any changes
        string backupFile = "input_backup.xlsx";
        workbook.Save(backupFile); // backup created

        // Access the collection of external links
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Iterate through each external link and modify its path as needed
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Example modification: replace an old folder path with a new one
            string oldPath = link.DataSource;
            string newPath = oldPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");
            link.DataSource = newPath;
        }

        // Save the workbook with updated external link paths
        string updatedFile = "output.xlsx";
        workbook.Save(updatedFile);
    }
}