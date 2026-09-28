// Title: How to backup an Excel workbook and update external link folder paths using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, creates a backup copy, then iterates through workbook.Worksheets.ExternalLinks to replace a given old folder path with a new one, updating both OriginalDataSource and DataSource. | Show how to construct a backup filename by appending "_backup" to the original name and save the workbook using Aspose.Cells before any modifications. | Provide a C# snippet that saves the workbook with the revised external link paths to a separate file after the replacements are applied.
// Common Searches: Aspose.Cells .NET backup workbook before editing external links | C# replace folder path in external links of an Excel file using Aspose.Cells | how to update OriginalDataSource and DataSource for external links with Aspose.Cells | save modified Excel workbook as new file after changing external link paths in C# | create backup copy of Excel workbook programmatically with Aspose.Cells
// Tags: Aspose.Cells backup workbook C# | replace external link folder path Aspose.Cells | sync external link data sources Aspose.Cells | save modified Excel file Aspose.Cells | external link path manipulation .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkBackupExample
{
    // The example loads input.xlsx, generates a backup file named input_backup.xlsx, iterates through all external links to replace an old base folder with a new one while synchronizing OriginalDataSource and DataSource, and finally saves the updated workbook as input_modified.xlsx.
    class Program
    {
        static void Main()
        {
            // Path to the original workbook
            string originalPath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(originalPath);

            // -----------------------------------------------------------------
            // Step 1: Create a backup of the workbook before any modifications
            // -----------------------------------------------------------------
            string backupPath = Path.Combine(
                Path.GetDirectoryName(originalPath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(originalPath) + "_backup" + Path.GetExtension(originalPath));

            // Save the backup (save rule)
            workbook.Save(backupPath);

            // ---------------------------------------------------------------
            // Step 2: Modify external link paths in the loaded workbook
            // ---------------------------------------------------------------
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Example: replace an old base folder with a new one in each link
            string oldBase = @"C:\OldExternalFolder\";
            string newBase = @"D:\NewExternalFolder\";

            for (int i = 0; i < externalLinks.Count; i++)
            {
                // Use OriginalDataSource to keep the original value if needed
                string currentSource = externalLinks[i].OriginalDataSource;

                // If OriginalDataSource is empty, fall back to DataSource
                if (string.IsNullOrEmpty(currentSource))
                {
                    currentSource = externalLinks[i].DataSource;
                }

                // Perform the path replacement
                string updatedSource = currentSource.Replace(oldBase, newBase, StringComparison.OrdinalIgnoreCase);

                // Apply the updated path
                externalLinks[i].OriginalDataSource = updatedSource;
                // Also update DataSource to keep both properties in sync
                externalLinks[i].DataSource = updatedSource;
            }

            // ---------------------------------------------------------------
            // Step 3: Save the workbook with modified external links
            // ---------------------------------------------------------------
            string modifiedPath = "input_modified.xlsx";
            workbook.Save(modifiedPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Backup created at: {backupPath}");
            Console.WriteLine($"Modified workbook saved at: {modifiedPath}");
        }
    }
}
