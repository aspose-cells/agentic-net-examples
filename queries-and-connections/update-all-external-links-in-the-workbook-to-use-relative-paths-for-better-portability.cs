using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class UpdateExternalLinksToRelativePaths
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook that contains external links
                string sourceFilePath = @"C:\Data\MainWorkbook.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(sourceFilePath);

                // Determine the directory of the workbook – this will be the base for relative paths
                string workbookDirectory = Path.GetDirectoryName(sourceFilePath);

                // Access the collection of external links
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                // Iterate through each external link and convert its DataSource to a relative path
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];

                    // Original data source (could be absolute path, UNC path, or URL)
                    string originalPath = link.DataSource;

                    // Only process if the path is rooted (i.e., an absolute file system path)
                    if (!string.IsNullOrEmpty(originalPath) && Path.IsPathRooted(originalPath))
                    {
                        // Compute the relative path from the workbook's folder to the external file
                        string relativePath = Path.GetRelativePath(workbookDirectory, originalPath);

                        // Update the external link to use the relative path
                        link.DataSource = relativePath;

                        // Optionally also update OriginalDataSource to keep consistency
                        link.OriginalDataSource = relativePath;
                    }
                }

                // Prepare output path and ensure its directory exists
                string outputFilePath = @"C:\Data\MainWorkbook_RelativeLinks.xlsx";
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputFilePath);

                Console.WriteLine("External links have been updated to relative paths and workbook saved to:");
                Console.WriteLine(outputFilePath);
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}