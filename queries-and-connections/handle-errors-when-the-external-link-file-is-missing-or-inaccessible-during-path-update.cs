using System;
using System.IO;
using Aspose.Cells;

class UpdateExternalLinksWithErrorHandling
{
    static void Main()
    {
        // Load the workbook that contains external links
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the collection of external links
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Iterate through each external link
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Current data source (path) of the external link
            string currentPath = link.DataSource;

            // Example transformation: replace an old base folder with a new one
            string updatedPath = currentPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");

            // Verify that the updated file actually exists
            if (File.Exists(updatedPath))
            {
                // If the file exists, update the link's data source
                link.DataSource = updatedPath;
                Console.WriteLine($"Link updated to: {updatedPath}");
            }
            else
            {
                // If the file is missing or inaccessible, handle the error
                Console.WriteLine($"Missing external file: {updatedPath}. Removing the link.");

                // Remove the problematic link to keep the workbook consistent
                externalLinks.RemoveAt(i);
                i--; // Adjust index because the collection size has decreased
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}