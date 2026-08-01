// Title: Update External Link Paths in Excel Workbooks with Aspose.Cells (C#)
// Description: Loads a workbook, scans its ExternalLinkCollection, replaces the old folder segment in DataSource and OriginalDataSource with a new directory, and saves the file so all linked formulas point to the relocated source workbooks.
// Keywords: Aspose.Cells external links | C# update Excel link paths | replace DataSource folder | OriginalDataSource path change | Excel workbook external reference migration
// Common Searches: change external link file path Aspose.Cells | update Excel workbook links after moving files | programmatically modify DataSource in .NET | Aspose.Cells external link path replacement example
// Developer Intent: Rewrite the file system paths of external links in an Excel workbook so they reference a new directory.
// Use Cases: Reorganize source files on a server and keep all linked workbooks functional. | Move a project between drives or network shares without breaking data connections. | Automate link updates during deployment to staging or production environments.
// AI Prompts: Write C# code using Aspose.Cells to batch‑update external link paths after relocating source workbooks. | Show how to validate that every ExternalLink.DataSource points to the new folder after saving. | Add robust error handling for missing or inaccessible external files when updating link paths.

using System;
using Aspose.Cells;

// Loads a workbook, scans its ExternalLinkCollection, replaces the old folder segment in DataSource and OriginalDataSource with a new directory, and saves the file so all linked formulas point to the relocated source workbooks.
class UpdateExternalLinks
{
    static void Main()
    {
        // Load the workbook that contains external links
        string sourceWorkbookPath = @"C:\OldFolder\MainWorkbook.xlsx";
        Workbook workbook = new Workbook(sourceWorkbookPath);

        // Define the old and new base directories for the external source files
        string oldBasePath = @"C:\OldFolder\ExternalSources\";
        string newBasePath = @"D:\NewFolder\ExternalSources\";

        // Get the collection of external links in the workbook
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Update each external link's path to point to the new directory
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Update the DataSource property if it contains the old base path
            if (!string.IsNullOrEmpty(link.DataSource) &&
                link.DataSource.StartsWith(oldBasePath, StringComparison.OrdinalIgnoreCase))
            {
                link.DataSource = link.DataSource.Replace(oldBasePath, newBasePath);
            }

            // Also update the OriginalDataSource property to keep both in sync
            if (!string.IsNullOrEmpty(link.OriginalDataSource) &&
                link.OriginalDataSource.StartsWith(oldBasePath, StringComparison.OrdinalIgnoreCase))
            {
                link.OriginalDataSource = link.OriginalDataSource.Replace(oldBasePath, newBasePath);
            }
        }

        // Save the workbook with updated external link paths
        string updatedWorkbookPath = @"D:\NewFolder\MainWorkbook_Updated.xlsx";
        workbook.Save(updatedWorkbookPath);
    }
}
