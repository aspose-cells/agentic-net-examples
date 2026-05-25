using System;
using Aspose.Cells;

class UpdateExternalLinkPaths
{
    static void Main()
    {
        // Path to the workbook that contains external links
        string workbookPath = @"C:\OldFolder\MainWorkbook.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Old and new base directories for the external source workbooks
        string oldBaseDir = @"C:\OldFolder\ExternalSources\";
        string newBaseDir = @"D:\NewFolder\ExternalSources\";

        // Get the collection of external links in the workbook
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Iterate through each external link and update its path
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Update the DataSource property if it contains the old base directory
            if (!string.IsNullOrEmpty(link.DataSource) &&
                link.DataSource.StartsWith(oldBaseDir, StringComparison.OrdinalIgnoreCase))
            {
                string updatedDataSource = link.DataSource.Replace(oldBaseDir, newBaseDir);
                link.DataSource = updatedDataSource; // uses ExternalLink.DataSource property
            }

            // Also update the OriginalDataSource property for completeness
            if (!string.IsNullOrEmpty(link.OriginalDataSource) &&
                link.OriginalDataSource.StartsWith(oldBaseDir, StringComparison.OrdinalIgnoreCase))
            {
                string updatedOriginal = link.OriginalDataSource.Replace(oldBaseDir, newBaseDir);
                link.OriginalDataSource = updatedOriginal; // uses ExternalLink.OriginalDataSource property
            }
        }

        // Save the updated workbook (overwrite or to a new location)
        string outputPath = @"C:\NewFolder\MainWorkbook_Updated.xlsx";
        workbook.Save(outputPath);
    }
}