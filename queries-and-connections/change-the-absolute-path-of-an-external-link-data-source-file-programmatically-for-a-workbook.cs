using System;
using Aspose.Cells;

class ChangeExternalLinkPath
{
    static void Main()
    {
        // Load the workbook that contains external links
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each external link and update its stored path
        for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
        {
            ExternalLink link = workbook.Worksheets.ExternalLinks[i];

            // Get the current absolute path of the data source
            string originalPath = link.OriginalDataSource;

            // Example replacement: change the root folder from C:\OldFolder\ to D:\NewFolder\
            string updatedPath = originalPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");

            // Assign the modified path back to the external link
            link.OriginalDataSource = updatedPath;
        }

        // Save the workbook with the updated external link paths
        workbook.Save("output.xlsx");
    }
}