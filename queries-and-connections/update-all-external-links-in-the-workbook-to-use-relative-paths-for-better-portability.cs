// Title: Update external links to relative paths in an Aspose.Cells workbook (C#)
// Description: Loads a workbook, determines its directory, iterates the Worksheets.ExternalLinks collection, converts any absolute DataSource path to a relative path using Path.GetRelativePath, and saves the file. Includes validation, error handling, and optional output location for portable Excel files.
// Keywords: Aspose.Cells external links | C# relative path Excel | convert absolute link to relative | workbook portability | .NET Excel external data source | Path.GetRelativePath Aspose | update external links programmatically
// Common Searches: Aspose.Cells change external link to relative C# | make Excel workbook portable by updating link paths | C# code to replace absolute DataSource in Aspose.Cells | how to use Path.GetRelativePath with Aspose external links | batch convert external links to relative in Excel files
// Developer Intent: Replace absolute external link paths with relative ones to ensure the workbook works after being moved or shared.
// Use Cases: Distribute a template workbook that references companion files without breaking links. | Integrate into CI/CD pipelines to standardize link paths before committing to source control. | Create a migration tool that updates legacy workbooks for cloud‑based storage environments.
// AI Prompts: Write C# code using Aspose.Cells that scans workbook.Worksheets.ExternalLinks and rewrites each absolute DataSource to a relative path based on the workbook's folder. | Provide a robust method that checks for null or empty DataSource values, handles exceptions, and logs which links were changed. | Explain how to test that the new relative paths resolve correctly after saving the workbook, including sample verification code.

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, determines its directory, iterates the Worksheets.ExternalLinks collection, converts any absolute DataSource path to a relative path using Path.GetRelativePath, and saves the file. Includes validation, error handling, and optional output location for portable Excel files.
class UpdateExternalLinksToRelative
{
    static void Main()
    {
        // Path to the source workbook that contains external links
        string sourceFile = @"C:\Data\MainWorkbook.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file not found: {sourceFile}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(sourceFile);

            // Determine the folder where the workbook resides – this will be the base for relative paths
            string workbookFolder = Path.GetDirectoryName(workbook.AbsolutePath ?? sourceFile);

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and replace its absolute path with a relative one
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // Only process if the current DataSource is an absolute path
                if (!string.IsNullOrEmpty(link.DataSource) && Path.IsPathRooted(link.DataSource))
                {
                    // Compute the relative path from the workbook folder to the external file
                    string relativePath = Path.GetRelativePath(workbookFolder, link.DataSource);

                    // Update the link to use the relative path
                    link.DataSource = relativePath;
                }
            }

            // Save the workbook (overwrites the original file or you can specify a new path)
            string outputFile = @"C:\Data\MainWorkbook_Relative.xlsx";
            workbook.Save(outputFile);

            Console.WriteLine("External links have been updated to relative paths and saved to:");
            Console.WriteLine(outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
