using System;
using Aspose.Cells;

class ExternalLinkAudit
{
    static void Main()
    {
        // Load the workbook that contains external links
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each external link in the workbook
        for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
        {
            ExternalLink link = workbook.Worksheets.ExternalLinks[i];

            // Log the original external link path
            string originalPath = link.OriginalDataSource;
            Console.WriteLine($"External Link {i} - Original Path: {originalPath}");

            // Example modification: replace an old base URL with a new one
            string updatedPath = originalPath.Replace(
                @"https://oldserver.com/",
                @"/new/shared/documents/");

            // Update the stored original data source with the new path
            link.OriginalDataSource = updatedPath;

            // Log the updated external link path
            Console.WriteLine($"External Link {i} - Updated Path: {link.OriginalDataSource}");
        }

        // Save the workbook after modifications
        workbook.Save("output.xlsx");
    }
}