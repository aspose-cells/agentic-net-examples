// Title: Update External Link Paths After Moving Source Workbooks – Aspose.Cells for .NET (C#)
// Description: C# example that loads a workbook, replaces the old folder segment in each external link's OriginalDataSource and DataSource, and saves the file with corrected paths using Aspose.Cells.
// Keywords: Aspose.Cells external link update | C# Excel external references | change workbook link path .NET | ExternalLinkCollection path replace | fix broken Excel links programmatically | batch update external links | GitHub Aspose.Cells example | Excel formula external source relocation
// Common Searches: Aspose.Cells change external link folder | Update Excel external references after moving files C# | Replace old directory in external links Aspose.Cells | Programmatically fix broken external links in Excel | C# code to update workbook external link paths
// Developer Intent: Modify the file paths of all external links in an Excel workbook so that formulas reference the new location of source workbooks.
// Use Cases: Repair broken external references after a server or folder migration. | Automate path correction for dozens of workbooks without opening Excel. | Prepare a workbook for distribution by ensuring all external links point to a standardized directory.
// AI Prompts: Generate C# code with Aspose.Cells that replaces a specific folder segment in every external link path of a workbook and saves the updated file. | Explain how to validate that external links were correctly updated after changing OriginalDataSource and DataSource properties. | Provide a C# method that logs each original and new external link path while performing the update with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that loads a workbook, replaces the old folder segment in each external link's OriginalDataSource and DataSource, and saves the file with corrected paths using Aspose.Cells.
class UpdateExternalLinks
{
    static void Main()
    {
        // Load the workbook that contains external links
        Workbook workbook = new Workbook("SourceWorkbook.xlsx");

        // Define the old directory and the new directory where the source workbooks were moved
        string oldDirectory = @"C:\OldFolder\";
        string newDirectory = @"D:\NewFolder\";

        // Get the collection of external links in the workbook
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Update each external link's path
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Replace the old directory part with the new directory in the original data source
            string updatedPath = link.OriginalDataSource.Replace(oldDirectory, newDirectory);

            // Apply the updated path to both OriginalDataSource and DataSource
            link.OriginalDataSource = updatedPath;
            link.DataSource = updatedPath;
        }

        // Save the workbook with the corrected external link paths
        workbook.Save("UpdatedWorkbook.xlsx");
    }
}
