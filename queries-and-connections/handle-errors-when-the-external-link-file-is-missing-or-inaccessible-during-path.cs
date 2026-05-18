using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkErrorHandling
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains external links
            Workbook workbook = new Workbook("MainWorkbook.xlsx");

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Define the part of the path that needs to be replaced (example)
            const string oldRoot = @"C:\OldExternalFiles\";
            const string newRoot = @"D:\NewExternalFiles\";

            // Iterate through the external links in reverse order to allow safe removal
            for (int i = externalLinks.Count - 1; i >= 0; i--)
            {
                ExternalLink link = externalLinks[i];
                string currentPath = link.DataSource; // current data source path

                // Check if the external file actually exists on disk
                if (!File.Exists(currentPath))
                {
                    // The file is missing or inaccessible – handle the error
                    Console.WriteLine($"External link not found: {currentPath}");
                    Console.WriteLine("Removing the invalid external link from the workbook.");

                    // Remove the invalid external link (uses RemoveAt rule)
                    externalLinks.RemoveAt(i);
                    continue;
                }

                // If the file exists, update its path if it matches the old root
                if (currentPath.StartsWith(oldRoot, StringComparison.OrdinalIgnoreCase))
                {
                    string updatedPath = currentPath.Replace(oldRoot, newRoot);
                    link.DataSource = updatedPath; // update the data source path
                    Console.WriteLine($"Updated external link path: {updatedPath}");
                }
            }

            // Save the workbook with the corrected external links
            workbook.Save("MainWorkbook_Updated.xlsx");
        }
    }
}