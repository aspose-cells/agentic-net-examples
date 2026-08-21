// Title: Change Excel External Link Paths to a UNC Network Share with Aspose.Cells for .NET
// Description: Load an XLS workbook, iterate through its Worksheets.ExternalLinks collection, replace the local folder segment in each link's OriginalDataSource with a UNC network share path, and save the modified workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# external link UNC | update Excel external link path | replace OriginalDataSource | network share Excel links | programmatic Excel link path change | C# workbook external links | UNC path conversion Aspose
// Common Searches: Aspose.Cells change external link path to UNC | C# update Excel external links network share | replace OriginalDataSource Aspose.Cells | modify external data source in XLS programmatically | convert local Excel link to network share
// Developer Intent: Programmatically replace the local file system path of each external link in an XLS workbook with a UNC network share path.
// Use Cases: Migrate legacy reports from a local server to a centralized file share without manual editing. | Automate deployment of Excel dashboards to a new network location for multi‑user access. | Prepare workbooks for cloud‑based file synchronization by converting absolute local paths to UNC paths.
// AI Prompts: Generate C# code using Aspose.Cells that loops through workbook.Worksheets.ExternalLinks, swaps a specified local folder prefix with a UNC folder, and saves the workbook. | Create a reusable method that accepts source workbook path, old local folder string, new UNC folder string, updates all ExternalLink OriginalDataSource values, and returns the output file path. | Explain how to handle external links that do not contain the old path prefix when performing the replacement with Aspose.Cells, ensuring no data loss.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkPathUpdate
{
    // Load an XLS workbook, iterate through its Worksheets.ExternalLinks collection, replace the local folder segment in each link's OriginalDataSource with a UNC network share path, and save the modified workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (local or existing location)
            string sourcePath = Path.GetFullPath("input.xls");

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all external links in the workbook
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                // Get the current external link
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                // OriginalDataSource holds the stored path of the external link
                string originalLink = link.OriginalDataSource;

                // Replace the local part of the path with the network shared folder UNC path
                // Example: replace "C:\\Data\\Reports\\" with "\\\\Server\\Shared\\Reports\\"
                string modifiedLink = originalLink.Replace(
                    @"C:\Data\Reports\",
                    @"\\Server\Shared\Reports\");

                // Assign the modified path back to the external link
                link.OriginalDataSource = modifiedLink;
            }

            // Save the workbook with the updated external link paths
            string outputPath = Path.Combine(Path.GetDirectoryName(sourcePath), "output.xls");
            workbook.Save(outputPath);
        }
    }
}
