using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkPathUpdate
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (XLS format)
            string sourcePath = Path.GetFullPath("input.xls");

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all external links in the workbook
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                // Get the stored original data source path
                string originalPath = link.OriginalDataSource;

                // Replace the local folder part with a network shared folder (UNC path)
                // Example: replace "C:\Temp\" with "\\Server\SharedFolder\"
                string modifiedPath = originalPath.Replace(
                    @"C:\Temp\",
                    @"\\Server\SharedFolder\");

                // Update the external link with the new path
                link.OriginalDataSource = modifiedPath;
            }

            // Save the modified workbook
            string outputPath = Path.Combine(Path.GetDirectoryName(sourcePath), "output.xls");
            workbook.Save(outputPath);
        }
    }
}