// Title: Convert Excel external links to relative paths with Aspose.Cells (C#)
// Description: A C# console app that loads an Excel workbook, scans all external links, resolves each link to an absolute path, computes a relative path using `Path.GetRelativePath`, updates both `DataSource` and `OriginalDataSource`, and saves the workbook. The result is a portable file that retains functional external references after being moved to another folder or system.
// Keywords: Aspose.Cells external links | C# relative path Excel | convert absolute link to relative | make workbook portable | Path.GetRelativePath Aspose | update DataSource OriginalDataSource | .NET Excel link management
// Common Searches: Aspose.Cells change external link to relative path | C# make Excel workbook portable | convert absolute external links in .xlsx to relative | update DataSource OriginalDataSource Aspose.Cells | Path.GetRelativePath example with Excel
// Developer Intent: Replace absolute external link paths with relative ones so the workbook remains functional after being moved across environments.
// Use Cases: Prepare a workbook for distribution by converting all external references to relative paths. | Automate link updates in a CI/CD pipeline to ensure Excel files work on any build agent. | Synchronize `OriginalDataSource` with `DataSource` after path conversion to keep metadata consistent.
// AI Prompts: Generate a C# method that takes input and output workbook paths and rewrites every external link to a relative path using Aspose.Cells. | Explain why both `DataSource` and `OriginalDataSource` should be updated when converting external links to relative paths. | Show a step‑by‑step tutorial on using `Path.GetRelativePath` with Aspose.Cells to make Excel external links portable.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkRelativePath
{
    // A C# console app that loads an Excel workbook, scans all external links, resolves each link to an absolute path, computes a relative path using `Path.GetRelativePath`, updates both `DataSource` and `OriginalDataSource`, and saves the workbook. The result is a portable file that retains functional external references after being moved to another folder or system.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and output workbook path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: AsposeCellsExternalLinkRelativePath <input.xlsx> <output.xlsx>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Directory of the workbook file – used as the base for relative paths
            string workbookDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath));

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and replace its DataSource with a relative path
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // Resolve the current DataSource to an absolute path (if it is already relative, combine with workbook directory)
                string absoluteLinkPath = Path.IsPathRooted(link.DataSource)
                    ? Path.GetFullPath(link.DataSource)
                    : Path.GetFullPath(Path.Combine(workbookDirectory, link.DataSource));

                // Compute the relative path from the workbook directory to the external file
                string relativePath = Path.GetRelativePath(workbookDirectory, absoluteLinkPath);

                // Update the link to use the relative path
                link.DataSource = relativePath;

                // Optionally also update OriginalDataSource to keep consistency
                link.OriginalDataSource = relativePath;
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved with relative external links to: {outputPath}");
        }
    }
}
