// Title: Update Excel External Link UNC Paths with Aspose.Cells for .NET (C#)
// Description: Learn how to load an Excel workbook, enumerate its ExternalLinkCollection, and replace old UNC share prefixes in the DataSource and OriginalDataSource properties with a new network share path. The example saves the modified workbook, ensuring all external formulas point to the correct location after a server migration.
// Keywords: Aspose.Cells external links | C# update UNC path | Excel external link DataSource | replace network share in workbook | modify OriginalDataSource Aspose | batch update Excel links .NET | external link collection Aspose.Cells
// Common Searches: change external link source UNC path Aspose.Cells | update Excel workbook external links C# | replace old server share with new share in Excel file | Aspose.Cells modify ExternalLinkCollection | programmatically edit DataSource of external links
// Developer Intent: Replace an outdated UNC network share prefix in every external link of an Excel workbook with a new share path using Aspose.Cells for .NET.
// Use Cases: Migrate workbooks after moving a data server so formulas continue to resolve. | Automate bulk updates of multiple Excel files before distribution. | Synchronize OriginalDataSource after a path change to keep audit trails accurate.
// AI Prompts: Write C# code that uses Aspose.Cells to scan all external links in an Excel workbook and replace a given UNC prefix with a new one, then save the file. | Explain how to verify that DataSource and OriginalDataSource values were correctly updated after changing the network share path with Aspose.Cells.

using System;
using Aspose.Cells;

// Learn how to load an Excel workbook, enumerate its ExternalLinkCollection, and replace old UNC share prefixes in the DataSource and OriginalDataSource properties with a new network share path. The example saves the modified workbook, ensuring all external formulas point to the correct location after a server migration.
class UpdateExternalLinks
{
    static void Main()
    {
        // Load the workbook that contains external links
        Workbook workbook = new Workbook("input.xlsx");

        // Define the old and new network share prefixes
        string oldShare = @"\\oldserver\share\";
        string newShare = @"\\newserver\share\";

        // Get the collection of external links
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Iterate through each external link and update its path
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Update DataSource if it uses the old share path
            if (!string.IsNullOrEmpty(link.DataSource) &&
                link.DataSource.StartsWith(oldShare, StringComparison.OrdinalIgnoreCase))
            {
                string updatedPath = newShare + link.DataSource.Substring(oldShare.Length);
                link.DataSource = updatedPath;
            }

            // Update OriginalDataSource similarly (optional but ensures consistency)
            if (!string.IsNullOrEmpty(link.OriginalDataSource) &&
                link.OriginalDataSource.StartsWith(oldShare, StringComparison.OrdinalIgnoreCase))
            {
                string updatedOriginal = newShare + link.OriginalDataSource.Substring(oldShare.Length);
                link.OriginalDataSource = updatedOriginal;
            }
        }

        // Save the workbook with updated external link URLs
        workbook.Save("output.xlsx");
    }
}
