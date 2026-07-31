// Title: Aspose.Cells C# – Convert absolute external links to relative paths for portable Excel workbooks
// Description: The sample loads an Excel file, iterates its ExternalLinkCollection, calculates a workbook‑based relative reference for each absolute DataSource using Path.GetRelativePath, updates the link, and saves the result, allowing the file to be moved without breaking external connections.
// Keywords: Aspose.Cells | C# external links | relative reference conversion | Excel workbook portability | ExternalLink.DataSource | Path.GetRelativePath | remove absolute paths | cross‑environment Excel | automate link update
// Common Searches: how to change external link paths to relative in Aspose.Cells .NET | make Excel workbook portable by updating external links with C# | convert absolute external link to relative path Aspose.Cells example | Aspose.Cells replace absolute DataSource with relative path | C# code for adjusting external links after moving workbook
// Developer Intent: Replace absolute DataSource values of external links with workbook‑relative references so the file remains functional after relocation.
// Use Cases: Shift a workbook that pulls data from CSV files to a new folder hierarchy without breaking the links. | Distribute a template to team members where each copy resolves external sources relative to its own location. | Automate the preparation of existing reports for deployment on a shared server that uses a different directory structure.
// AI Prompts: Write C# code using Aspose.Cells that scans all external links in a workbook and rewrites each DataSource to a relative reference based on the workbook’s directory. | Explain the interaction between Path.GetRelativePath and ExternalLink.DataSource, highlighting edge cases such as missing files or UNC paths. | Provide an enhanced version of the sample that logs every modified link, skips links that are already relative, and handles errors gracefully.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkRelativePath
{
    // The sample loads an Excel file, iterates its ExternalLinkCollection, calculates a workbook‑based relative reference for each absolute DataSource using Path.GetRelativePath, updates the link, and saves the result, allowing the file to be moved without breaking external connections.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output workbook paths (adjust as needed)
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Directory of the workbook – used as the base for relative paths
            string workbookDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath));

            // Get the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and replace its DataSource with a relative path
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // Full (absolute) path of the current external link
                string absoluteLinkPath = link.DataSource;

                // If the link is already a relative path, skip processing
                if (!Path.IsPathRooted(absoluteLinkPath))
                    continue;

                // Compute the relative path from the workbook's directory to the external file
                string relativePath = Path.GetRelativePath(workbookDirectory, absoluteLinkPath);

                // Update the external link to use the relative path
                link.DataSource = relativePath;
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
        }
    }
}
