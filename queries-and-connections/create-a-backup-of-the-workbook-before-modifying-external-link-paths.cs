// Title: Backup Excel Workbook and Update External Link Paths with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, saves a backup copy, iterates through the ExternalLinkCollection to replace a folder segment in each link's OriginalDataSource, then overwrites the original file with the updated links and releases resources.
// Keywords: Aspose.Cells | C# backup Excel workbook | external link collection | modify OriginalDataSource | replace folder path | save workbook copy | Excel external links .NET | Aspose.Cells backup example
// Common Searches: How to backup an Excel file before editing external links using Aspose.Cells | C# code to change folder name in external link paths | Aspose.Cells create workbook copy then update external links | Backup and modify external data sources in Excel with .NET | Save a backup of a workbook before changing external links
// Developer Intent: Create a backup of the workbook, then modify each external link's OriginalDataSource and save the changes back to the original file.
// Use Cases: Automated migration of data‑source folders across many workbooks while preserving the originals. | Safety net for batch scripts that rewrite external links in deployment pipelines. | Audit‑trail generation by keeping an unchanged copy before correcting broken links. | Disaster‑recovery scenario where the original workbook must remain intact during path updates.
// AI Prompts: Generate C# code using Aspose.Cells that copies a workbook to a backup file, iterates through workbook.Worksheets.ExternalLinks, replaces a specified old folder name with a new one in OriginalDataSource, and saves the changes back to the original workbook. | Provide a reusable method accepting (string sourcePath, string backupPath, string oldFolder, string newFolder) that creates a backup, updates all external link paths, and returns a success status. | Explain how to safely update external link paths in an Excel file with Aspose.Cells while keeping a versioned backup.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackupExample
{
    // Loads an existing workbook, saves a backup copy, iterates through the ExternalLinkCollection to replace a folder segment in each link's OriginalDataSource, then overwrites the original file with the updated links and releases resources.
    class Program
    {
        static void Main()
        {
            // Path to the original workbook
            string originalPath = "OriginalWorkbook.xlsx";

            // Path for the backup copy
            string backupPath = "OriginalWorkbook_Backup.xlsx";

            // Load the original workbook (uses the provided load rule)
            Workbook workbook = new Workbook(originalPath);

            // Create a backup before making any changes (uses the provided save rule)
            workbook.Save(backupPath);

            // Modify external link paths
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                // Example modification: replace a folder segment in the original data source
                string original = externalLinks[i].OriginalDataSource;
                if (!string.IsNullOrEmpty(original))
                {
                    // Adjust the path as needed; here we replace "OldFolder" with "NewFolder"
                    string modified = original.Replace("OldFolder", "NewFolder");
                    externalLinks[i].OriginalDataSource = modified;
                }
            }

            // Save the modified workbook (uses the provided save rule)
            workbook.Save(originalPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Backup created at: " + Path.GetFullPath(backupPath));
            Console.WriteLine("External links updated and original workbook saved.");
        }
    }
}
